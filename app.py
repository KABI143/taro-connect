from flask import Flask, render_template, request, session
import pandas as pd
import io
import os
import gc
from flask import jsonify
import tempfile

app = Flask(__name__)
app.secret_key = "taroconnect-secret-key-2024"

import threading
_progress_lock = threading.Lock()
progress_data = {"percent": 0, "status": "Idle"}

def set_progress(percent, status):
    with _progress_lock:
        progress_data["percent"] = percent
        progress_data["status"]  = status

# ═══════════════════════════════════════════════════════
# CONSTANTS
# ═══════════════════════════════════════════════════════

ERROR_CODES = {
    0:      "No Error",
    1:      "Tank Full / Sump Dry",
    2:      "Low Voltage",
    4:      "High Voltage",
    8:      "Short Circuit",
    16:     "Overload",
    32:     "Dry Run",
    64:     "Pressure Cut In",
    128:    "Pressure Cut Off",
    256:    "Flow Min",
    512:    "Flow Max",
    1024:   "Pressure Sensor Not Connected",
    2048:   "Flow Sensor Not Connected",
    4096:   "Phase Sequence Fail",
    8192:   "Y Phase Fail",
    16384:  "Power Failure",
    32768:  "Phase Unbalance",
    65536:  "Current Unbalance",
    131072: "Pressure Sensor Fault",
    262144: "R Phase Fail",
    524288: "B Phase Fail",
    786432: "R/B Phase Fail",
}

MODE_DESCRIPTIONS = {0: "Manual Mode", 1: "Auto Mode", 2: "Timer Mode",
                     3: "Schedule Mode", 4: "Bypass Mode"}
MODE_ICONS        = {0: "🔧", 1: "🤖", 2: "⏱", 3: "📅", 4: "⚡"}

CRITICAL_ERRORS = {8, 16}          # Short Circuit, Overload only
WARNING_ERRORS  = {1, 2, 4, 32, 64, 128, 256, 512, 1024, 2048, 4096, 8192,
                   16384, 32768, 65536, 131072, 262144, 524288, 786432}

MAX_FILE_SIZE_MB   = 20
ALLOWED_EXTENSIONS = {".xlsx", ".xls"}

VOLTAGE_COLS    = ["Line Voltage", "Line Voltage 2", "Line Voltage 3"]
CURRENT_COLS    = ["Current Amp", "Current Amp2", "Current Amp3"]
KEY_SENSOR_COLS = VOLTAGE_COLS + CURRENT_COLS + ["Pressure", "Flow Sensor", "Frequency"]

OFFLINE_GAP_MINUTES = 3   # gaps larger than this = device offline

STATS_COLS = [
    "DeviceId", "IoTHubName", "PumpPhaseType", "QueuedTime-IST",
    "MotorRunningStatus", "Error CondMon", "ModeOfOperating",
    "Total Running Time", "PackCount", "NetType",
] + KEY_SENSOR_COLS + ["Signal"]

CHART_COLS = ["QueuedTime-IST"] + KEY_SENSOR_COLS + ["Signal", "PackCount", "NetType",
              "MotorRunningStatus", "Error CondMon"]

#CHART_DOWNSAMPLE = 500
CHART_DOWNSAMPLE = 999999 

# ═══════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════

def validate_upload(file):
    if not file or file.filename == "":
        return False, "No file selected."
    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in ALLOWED_EXTENSIONS:
        return False, f"Unsupported format '{ext}'. Only .xlsx / .xls allowed."
    file.seek(0, 2)
    size_mb = file.tell() / (1024 * 1024)
    file.seek(0)
    if size_mb > MAX_FILE_SIZE_MB:
        return False, f"File size {size_mb:.1f} MB exceeds {MAX_FILE_SIZE_MB} MB limit."
    return True, "OK"


def read_excel_cols(buf, engine, wanted_cols):
    buf.seek(0)
    header = pd.read_excel(buf, engine=engine, nrows=0)
    use    = [c for c in wanted_cols if c in header.columns]
    del header
    buf.seek(0)
    return pd.read_excel(buf, engine=engine, usecols=use)


def col_stats(series):
    s = pd.to_numeric(series, errors="coerce").dropna()
    if s.empty:
        return 0, 0, 0
    return round(float(s.min()), 2), round(float(s.max()), 2), round(float(s.mean()), 2)


def col_stats_motor_on(df, col):
    """Compute min/max/avg for a column but ONLY when motor is running.
    Current and voltage are 0 when motor is OFF — including those rows gives wrong averages."""
    if col not in df.columns or "MotorRunningStatus" not in df.columns:
        return 0, 0, 0
    motor_on = df["MotorRunningStatus"].astype(str).str.strip().str.upper().isin(
        ["1", "TRUE", "YES", "ON"])
    s = pd.to_numeric(df.loc[motor_on, col], errors="coerce").dropna()
    if s.empty:
        return 0, 0, 0
    return round(float(s.min()), 2), round(float(s.max()), 2), round(float(s.mean()), 2)


def group_avg(df, cols):
    present = [c for c in cols if c in df.columns]
    if not present:
        return 0
    vals = [pd.to_numeric(df[c], errors="coerce").mean() for c in present]
    vals = [v for v in vals if not pd.isna(v)]
    if not vals:
        return 0
    return round(sum(vals) / len(vals), 2)


def group_avg_motor_on(df, cols):
    """Average across phases but ONLY for motor-ON rows.
    Motor-OFF rows have 0 current/voltage which skews the average badly."""
    if "MotorRunningStatus" not in df.columns:
        return group_avg(df, cols)
    motor_on = df["MotorRunningStatus"].astype(str).str.strip().str.upper().isin(
        ["1", "TRUE", "YES", "ON"])
    present = [c for c in cols if c in df.columns]
    if not present:
        return 0
    vals = [pd.to_numeric(df.loc[motor_on, c], errors="coerce").mean() for c in present]
    vals = [v for v in vals if not pd.isna(v)]
    if not vals:
        return 0
    return round(sum(vals) / len(vals), 2)


def calc_avg_voltage_no_fail(df, cols):
    """Calculate average voltage excluding phase fail and power fail error codes.
    Phase fail codes: 4096, 8192, 262144, 524288, 786432
    Power failure code: 16384
    """
    phase_fail_codes = {4096, 8192, 262144, 524288, 786432}
    power_fail_code = 16384
    fail_codes = phase_fail_codes | {power_fail_code}
    
    if "Error CondMon" not in df.columns:
        return group_avg(df, cols)
    
    error_col = pd.to_numeric(df["Error CondMon"], errors="coerce").fillna(0).astype(int)
    no_fail_mask = ~error_col.isin(fail_codes)
    
    present = [c for c in cols if c in df.columns]
    if not present:
        return 0
    
    vals = [pd.to_numeric(df.loc[no_fail_mask, c], errors="coerce").mean() for c in present]
    vals = [v for v in vals if not pd.isna(v)]
    if not vals:
        return 0
    return round(sum(vals) / len(vals), 2)


def col_stats_no_fail(df, col):
    """Compute min/max/avg for a voltage column, excluding rows where a phase-fail
    or power-fail error is active. Readings during those faults (0V / spikes) are
    not representative of normal supply conditions.
    Phase fail codes: 4096, 8192, 262144, 524288, 786432
    Power failure code: 16384
    """
    phase_fail_codes = {4096, 8192, 262144, 524288, 786432}
    power_fail_code = 16384
    fail_codes = phase_fail_codes | {power_fail_code}

    if col not in df.columns:
        return 0, 0, 0

    if "Error CondMon" in df.columns:
        error_col = pd.to_numeric(df["Error CondMon"], errors="coerce").fillna(0).astype(int)
        no_fail_mask = ~error_col.isin(fail_codes)
        s = pd.to_numeric(df.loc[no_fail_mask, col], errors="coerce").dropna()
    else:
        s = pd.to_numeric(df[col], errors="coerce").dropna()

    if s.empty:
        return 0, 0, 0
    return round(float(s.min()), 2), round(float(s.max()), 2), round(float(s.mean()), 2)


def calc_avg_current_motor_on(df, cols):
    """Calculate average current only when motor is running.
    Current is 0 when motor is OFF, which skews the average."""
    return group_avg_motor_on(df, cols)


def minutes_to_dhm(total_minutes):
    total_minutes = int(total_minutes)
    d = total_minutes // 1440
    h = (total_minutes % 1440) // 60
    m = total_minutes % 60
    return f"{d}D : {h:02d}H : {m:02d}M"


def downsample_series(series, max_pts=CHART_DOWNSAMPLE):
    n = len(series)
    if n <= max_pts:
        return pd.to_numeric(series, errors="coerce").fillna(0).tolist()
    step = n // max_pts
    return pd.to_numeric(series.iloc[::step].iloc[:max_pts], errors="coerce").fillna(0).tolist()


def downsample_list(lst, max_pts=CHART_DOWNSAMPLE):
    if len(lst) <= max_pts:
        return lst
    step = len(lst) // max_pts
    return lst[::step][:max_pts]


# ═══════════════════════════════════════════════════════
# DURATION CALCULATIONS (single-pass, correct)
# ═══════════════════════════════════════════════════════

def compute_durations(df):
    """
    Returns a dict with all duration breakdowns computed from
    consecutive-row time differences.
    Keys:
        total_log_mins          — full log span (excl offline gaps)
        motor_running_mins      — motor ON, any error state
        motor_idle_mins         — motor OFF, device online
        no_error_running_mins   — motor ON, error == 0
        error_running_mins      — motor ON, error != 0
        offline_mins            — gaps > OFFLINE_GAP_MINUTES
        error_duration_by_code  — {error_code: minutes}
        mode_runtime            — {mode_code: minutes}  (motor ON only)
    """
    result = dict(
        total_log_mins=0,
        motor_running_mins=0,
        motor_idle_mins=0,
        no_error_running_mins=0,
        error_running_mins=0,
        offline_mins=0,
        error_duration_by_code={},
        mode_runtime={},
    )

    needs = {"QueuedTime-IST", "MotorRunningStatus", "Error CondMon", "ModeOfOperating"}
    if not needs.issubset(df.columns):
        return result

    ts      = pd.to_datetime(df["QueuedTime-IST"], errors="coerce")
    motor   = df["MotorRunningStatus"].astype(str).str.strip().str.upper().isin(
                  ["1", "TRUE", "YES", "ON"])
    err     = pd.to_numeric(df["Error CondMon"], errors="coerce").fillna(0).astype(int)
    mode_s  = pd.to_numeric(df["ModeOfOperating"], errors="coerce").fillna(-1).astype(int)

    for i in range(len(df) - 1):
        t1, t2 = ts.iloc[i], ts.iloc[i + 1]
        if pd.isna(t1) or pd.isna(t2):
            continue
        diff = (t2 - t1).total_seconds() / 60
        if diff <= 0:
            continue

        if diff > OFFLINE_GAP_MINUTES:
            result["offline_mins"] += diff
            continue

        result["total_log_mins"] += diff
        is_running = bool(motor.iloc[i])
        err_code   = int(err.iloc[i])
        mode_code  = int(mode_s.iloc[i])

        # FIX: Track error duration regardless of motor ON/OFF state.
        # Errors like High Voltage, Phase Fail, Power Failure occur even when
        # the motor is OFF — the old logic only counted inside `if is_running`
        # which caused 0D:00H:00M for every error in this dataset.
        if err_code != 0:
            ec = result["error_duration_by_code"]
            ec[err_code] = ec.get(err_code, 0) + diff

        if is_running:
            result["motor_running_mins"] += diff
            if err_code != 0:
                result["error_running_mins"] += diff
            else:
                result["no_error_running_mins"] += diff
            mr = result["mode_runtime"]
            mr[mode_code] = mr.get(mode_code, 0) + diff
        else:
            result["motor_idle_mins"] += diff

    return result


# ═══════════════════════════════════════════════════════
# PASS 1 — compute all scalars
# ═══════════════════════════════════════════════════════

def compute_stats(df):
    total_rows = len(df)

    device_id  = str(df["DeviceId"].iloc[0])      if "DeviceId"      in df.columns else "—"
    iot_hub    = str(df["IoTHubName"].iloc[0])     if "IoTHubName"    in df.columns else "—"
    _raw       = str(df["PumpPhaseType"].iloc[0])  if "PumpPhaseType" in df.columns else "—"
    pump_phase = ("Three Phase Pump" if _raw == "1" else
                  "Single Phase Pump" if _raw == "0" else _raw)

    date_range = "—"
    if "QueuedTime-IST" in df.columns:
        ts = pd.to_datetime(df["QueuedTime-IST"], errors="coerce").dropna()
        if not ts.empty:
            date_range = f"{ts.min().strftime('%d %b %Y')} → {ts.max().strftime('%d %b %Y')}"

    # Motor start count
    start_count = 0
    if "MotorRunningStatus" in df.columns:
        raw = df["MotorRunningStatus"].fillna(0).astype(str).str.strip().str.upper()
        ms  = raw.isin({"1", "TRUE", "YES", "ON"})
        start_count = int((~ms.shift(1, fill_value=False) & ms).sum())

    # Error summary (count-based)
    error_summary = {}
    error_count_by_code = {}
    if "Error CondMon" in df.columns:
        codes = pd.to_numeric(df["Error CondMon"], errors="coerce").dropna().astype(int)
        for code, count in codes.value_counts().items():
            desc     = ERROR_CODES.get(code, "Unknown Error")
            severity = ("critical" if code in CRITICAL_ERRORS else
                        "warning"  if code in WARNING_ERRORS  else "info")
            error_summary[f"{code} - {desc}"] = {"count": int(count), "severity": severity}
            error_count_by_code[code] = int(count)

    # ── Duration breakdown (single pass) ───────────────
    dur = compute_durations(df)

    total_log_mins        = dur["total_log_mins"] + dur["offline_mins"]
    motor_running_mins    = dur["motor_running_mins"]
    motor_idle_mins       = dur["motor_idle_mins"]
    #error_running_mins    = dur["error_running_mins"]
    error_running_mins = sum(dur["error_duration_by_code"].values())
    no_error_running_mins = dur["no_error_running_mins"]
    offline_mins          = dur["offline_mins"]
    error_duration_by_code = dur["error_duration_by_code"]

    # ── Mode summary (with correct duration) ───────────
    mode_summary = {}
    if "ModeOfOperating" in df.columns:
        for code, count in df["ModeOfOperating"].value_counts().items():
            try:    code_int = int(code)
            except: code_int = -1
            mode_summary[
                MODE_DESCRIPTIONS.get(code_int, f"Unknown Mode ({code})")
            ] = {
                "code": code_int,
                "count": int(count),
                "icon": MODE_ICONS.get(code_int, "❓"),
                "duration": minutes_to_dhm(dur["mode_runtime"].get(code_int, 0)),
            }

    # ── Error duration table (for new section) ─────────
    error_duration_table = []
    for code, mins in sorted(error_duration_by_code.items()):
        error_duration_table.append({
            "code": code,
            "desc": ERROR_CODES.get(code, "Unknown Error"),
            "severity": ("critical" if code in CRITICAL_ERRORS else
                         "warning"  if code in WARNING_ERRORS  else "info"),
            "count": error_count_by_code.get(code, 0),
            "duration": minutes_to_dhm(mins),
            "minutes": round(mins, 1),
        })

    def _s(col): return df[col] if col in df.columns else pd.Series(dtype=float)

    # Voltage stats per phase — ALL rows included (errors, phase fail, power fail)
    # Chart shows all data including fault rows, so KPI stats match chart data
    v1_min, v1_max, v1_avg = col_stats(_s("Line Voltage"))
    v2_min, v2_max, v2_avg = col_stats(_s("Line Voltage 2"))
    v3_min, v3_max, v3_avg = col_stats(_s("Line Voltage 3"))
    avg_voltage             = group_avg(df, VOLTAGE_COLS)

    # Avg Voltage KPI card — phase fail + power fail rows EXCLUDED
    # (0V / spike readings during faults skew the overall average badly)
    avg_voltage_no_fail     = calc_avg_voltage_no_fail(df, VOLTAGE_COLS)

    # Current stats per phase — motor ON rows only
    # Current is 0 when motor is OFF — averaging all rows gives completely wrong (low) values
    c1_min, c1_max, c1_avg = col_stats_motor_on(df, "Current Amp")
    c2_min, c2_max, c2_avg = col_stats_motor_on(df, "Current Amp2")
    c3_min, c3_max, c3_avg = col_stats_motor_on(df, "Current Amp3")
    avg_current             = group_avg_motor_on(df, CURRENT_COLS)
    
    # Kept for backward compatibility — same value as avg_current now
    avg_current_motor_on_val = avg_current

    p_min,    p_max,    p_avg    = col_stats(_s("Pressure"))
    f_min,    f_max,    f_avg    = col_stats(_s("Flow Sensor"))
    freq_min, freq_max, freq_avg = col_stats(_s("Frequency"))
    sig_min,  sig_max,  sig_avg  = col_stats(_s("Signal"))

    pack_count_total = 0
    if "PackCount" in df.columns:
        diff2 = pd.to_numeric(df["PackCount"], errors="coerce").diff().fillna(0)
        pack_count_total = int(diff2[diff2 > 0].sum())

    net_4g_count = net_2g_count = 0
    if "NetType" in df.columns:
        is_4g        = df["NetType"].astype(str).str.upper() == "4G"
        net_4g_count = int(is_4g.sum())
        net_2g_count = len(df) - net_4g_count
    network_pct_4g = round(net_4g_count / max(net_4g_count + net_2g_count, 1) * 100, 1)

    # Health score
    score = 100.0
    for key, val in error_summary.items():
        try:    code = int(key.split(" - ")[0])
        except: continue
        if code == 0: continue
        cnt = val["count"]
        if   code in CRITICAL_ERRORS: score -= min(15, cnt * 3)
        elif code in WARNING_ERRORS:  score -= min(8,  cnt)
        else:                          score -= min(4,  cnt * 0.5)
    if avg_voltage > 0:
        score -= min(20, abs(avg_voltage - 415) / 415 * 60)
    present_c = [c for c in CURRENT_COLS if c in df.columns]
    if len(present_c) == 3:
        means   = [pd.to_numeric(df[c], errors="coerce").mean() for c in present_c]
        overall = sum(means) / 3
        if overall > 0:
            score -= min(20, max(abs(m - overall) / overall for m in means) * 40)
    present_k = sum(1 for c in KEY_SENSOR_COLS + ["Signal"] if c in df.columns)
    score -= (1 - present_k / (len(KEY_SENSOR_COLS) + 1)) * 20
    health_score = max(0, min(100, round(score, 1)))

    # Data quality
    quality_findings = []
    checks = [
        ("Line Voltage","R-Phase Voltage"), ("Line Voltage 2","Y-Phase Voltage"),
        ("Line Voltage 3","B-Phase Voltage"), ("Current Amp","R-Phase Current"),
        ("Current Amp2","Y-Phase Current"),  ("Current Amp3","B-Phase Current"),
        ("Pressure","Pressure"), ("Flow Sensor","Flow Sensor"),
        ("Frequency","Frequency"), ("Signal","Signal Strength"),
    ]
    for col, label in checks:
        if col not in df.columns:
            quality_findings.append({"field": label, "status": "missing",
                                     "message": "Column not found in dataset", "pct": 0})
            continue
        null_pct = round(df[col].isna().sum() / total_rows * 100, 1) if total_rows else 0
        zero_pct = round((df[col].fillna(0) == 0).sum() / total_rows * 100, 1) if total_rows else 0
        if null_pct > 10:
            quality_findings.append({"field": label, "status": "warn",
                                     "message": f"{null_pct}% missing values", "pct": 100 - null_pct})
        elif zero_pct > 50:
            quality_findings.append({"field": label, "status": "warn",
                                     "message": f"{zero_pct}% zero values", "pct": 100 - zero_pct})
        else:
            quality_findings.append({"field": label, "status": "ok",
                                     "message": f"Good — {100 - null_pct}% complete",
                                     "pct": 100 - null_pct})

    return dict(
        total_rows=total_rows, device_id=device_id, iot_hub=iot_hub,
        pump_phase=pump_phase, date_range=date_range, start_count=start_count,
        error_summary=error_summary, mode_summary=mode_summary,
        error_duration_table=error_duration_table,
        # Voltage phase stats
        v1_min=v1_min, v1_max=v1_max, v1_avg=v1_avg,
        v2_min=v2_min, v2_max=v2_max, v2_avg=v2_avg,
        v3_min=v3_min, v3_max=v3_max, v3_avg=v3_avg,
        avg_voltage=avg_voltage,
        avg_voltage_no_fail=avg_voltage_no_fail,  # NEW: Avg voltage excluding phase/power fail
        # Current phase stats
        c1_min=c1_min, c1_max=c1_max, c1_avg=c1_avg,
        c2_min=c2_min, c2_max=c2_max, c2_avg=c2_avg,
        c3_min=c3_min, c3_max=c3_max, c3_avg=c3_avg,
        avg_current=avg_current,
        avg_current_motor_on=avg_current_motor_on_val,  # NEW: Avg current when motor running
        # Process params
        p_min=p_min,    p_max=p_max,    p_avg=p_avg,
        f_min=f_min,    f_max=f_max,    f_avg=f_avg,
        freq_min=freq_min, freq_max=freq_max, freq_avg=freq_avg,
        sig_min=sig_min, sig_max=sig_max, sig_avg=sig_avg,
        # Network / pack
        pack_count_total=pack_count_total, net_4g_count=net_4g_count,
        net_2g_count=net_2g_count, network_pct_4g=network_pct_4g,
        # Duration breakdown
        total_log_mins=total_log_mins,
        motor_running_mins=motor_running_mins,
        motor_idle_mins=motor_idle_mins,
        error_running_mins=error_running_mins,
        no_error_running_mins=no_error_running_mins,
        offline_mins=offline_mins,
        error_duration_by_code=error_duration_by_code,
        # Health
        health_score=health_score,
        quality_findings=quality_findings,
    )


# ═══════════════════════════════════════════════════════
# PASS 2 — chart data only (downsampled)
# ═══════════════════════════════════════════════════════

def compute_charts(df):
    def _ds(col):
        if col not in df.columns: return []
        return downsample_series(df[col])

    def _has_data(*cols):
        for col in cols:
            if col not in df.columns: continue
            s = pd.to_numeric(df[col], errors="coerce").dropna()
            if len(s) > 0 and s.abs().sum() > 0: return True
        return False

    time_labels = []
    if "QueuedTime-IST" in df.columns:
        time_labels = downsample_list(df["QueuedTime-IST"].astype(str).tolist())

    net_numeric = []
    if "NetType" in df.columns:
        net_numeric = downsample_list(
            (df["NetType"].astype(str).str.upper() == "4G").map({True: 4, False: 2}).tolist()
        )

    # Motor status for overlay (0/1)
    motor_status = []
    if "MotorRunningStatus" in df.columns:
        ms = df["MotorRunningStatus"].astype(str).str.strip().str.upper().isin(
                 ["1","TRUE","YES","ON"]).map({True:1,False:0})
        motor_status = downsample_list(ms.tolist())

    # Error code series (non-zero = error)
    error_series = []
    if "Error CondMon" in df.columns:
        ec = pd.to_numeric(df["Error CondMon"], errors="coerce").fillna(0).astype(int)
        error_series = downsample_list(ec.tolist())

    return dict(
        labels=time_labels,
        voltage1=_ds("Line Voltage"),   voltage2=_ds("Line Voltage 2"), voltage3=_ds("Line Voltage 3"),
        current1=_ds("Current Amp"),    current2=_ds("Current Amp2"),   current3=_ds("Current Amp3"),
        pressure=_ds("Pressure"),       flow=_ds("Flow Sensor"),
        frequency=_ds("Frequency"),     signal=_ds("Signal"),
        pack_count_graph=_ds("PackCount"),
        network_numeric=net_numeric,
        motor_status=motor_status,
        error_series=error_series,
        has_voltage  =_has_data("Line Voltage", "Line Voltage 2", "Line Voltage 3"),
        has_current  =_has_data("Current Amp",  "Current Amp2",   "Current Amp3"),
        has_pressure =_has_data("Pressure"),
        has_flow     =_has_data("Flow Sensor"),
        has_frequency=_has_data("Frequency"),
        has_signal   =_has_data("Signal"),
        has_pack     =_has_data("PackCount"),
        has_network  ="NetType" in df.columns and len(net_numeric) > 0,
    )


# ═══════════════════════════════════════════════════════
# ROUTES
# ═══════════════════════════════════════════════════════

@app.route("/", methods=["GET"])
def home():
    return render_template("index.html")

@app.route("/progress")
def progress():
    return jsonify(progress_data)

@app.route("/", methods=["POST"])
def upload():
    file = request.files.get("file")
    ok, msg = validate_upload(file)
    if not ok:
        return render_template("index.html", upload_error=msg)

    filename = file.filename
    set_progress(5, "Uploading File...")
    try:
        ext    = os.path.splitext(filename)[1].lower()
        engine = "xlrd" if ext == ".xls" else "openpyxl"
        file_bytes = file.read()
        del file
    except Exception as e:
        set_progress(0, "Idle")
        return render_template("index.html", upload_error=f"Could not read file: {e}")

    set_progress(15, "Reading date range from Excel...")
    try:
        buf = io.BytesIO(file_bytes)
        buf.seek(0)
        header = pd.read_excel(buf, engine=engine, nrows=0)
        if "QueuedTime-IST" in header.columns:
            buf.seek(0)
            date_df = pd.read_excel(buf, engine=engine, usecols=["QueuedTime-IST"])
            ts = pd.to_datetime(date_df["QueuedTime-IST"], errors="coerce").dropna()
            del date_df
            min_date   = ts.min().strftime('%Y-%m-%d') if not ts.empty else None
            max_date   = ts.max().strftime('%Y-%m-%d') if not ts.empty else None
            total_rows = len(ts)
        else:
            min_date = max_date = None
            total_rows = 0
        del buf
        gc.collect()
    except Exception as e:
        min_date = max_date = None
        total_rows = 0

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(filename)[1])
    tmp.write(file_bytes)
    tmp.close()
    del file_bytes

    session["tmp_path"]   = tmp.name
    session["filename"]   = filename
    session["engine"]     = engine
    session["min_date"]   = min_date
    session["max_date"]   = max_date
    session["total_rows"] = total_rows

    set_progress(25, "File ready — select date range...")
    return render_template("date_range.html", filename=filename,
                           min_date=min_date, max_date=max_date, total_rows=total_rows)


@app.route("/r", methods=["GET", "POST"])
def process():
    if request.method == "GET":
        from flask import redirect
        return redirect("/")

    tmp_path = session.get("tmp_path")
    filename = session.get("filename", "report.xlsx")
    engine   = session.get("engine", "openpyxl")
    min_date = session.get("min_date")
    max_date = session.get("max_date")

    if not tmp_path or not os.path.exists(tmp_path):
        return render_template("index.html", upload_error="Session expired. Please upload again.")

    use_filter = request.form.get("filter_type") == "range"
    date_from  = request.form.get("date_from", "")
    date_to    = request.form.get("date_to", "")

    set_progress(30, "Reading Excel...")
    try:
        with open(tmp_path, "rb") as f:
            buf = io.BytesIO(f.read())
    except Exception as e:
        set_progress(0, "Idle")
        return render_template("index.html", upload_error=f"Could not read temp file: {e}")

    # ── PASS 1 ─────────────────────────────────────────
    try:
        df1 = read_excel_cols(buf, engine, STATS_COLS)
        if use_filter and date_from and date_to and "QueuedTime-IST" in df1.columns:
            df1["QueuedTime-IST"] = pd.to_datetime(df1["QueuedTime-IST"], errors="coerce")
            mask = ((df1["QueuedTime-IST"] >= pd.Timestamp(date_from)) &
                    (df1["QueuedTime-IST"] <= pd.Timestamp(date_to) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)))
            df1 = df1[mask].reset_index(drop=True)

        set_progress(50, "Calculating Statistics...")
        stats = compute_stats(df1)
        del df1
        gc.collect()
    except Exception as e:
        set_progress(0, "Idle")
        return render_template("index.html", upload_error=f"Could not process file: {e}")

    set_progress(70, "Generating Charts...")

    # ── PASS 2 ─────────────────────────────────────────
    try:
        df2 = read_excel_cols(buf, engine, CHART_COLS)
        if use_filter and date_from and date_to and "QueuedTime-IST" in df2.columns:
            df2["QueuedTime-IST"] = pd.to_datetime(df2["QueuedTime-IST"], errors="coerce")
            mask = ((df2["QueuedTime-IST"] >= pd.Timestamp(date_from)) &
                    (df2["QueuedTime-IST"] <= pd.Timestamp(date_to) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)))
            df2 = df2[mask].reset_index(drop=True)

        charts = compute_charts(df2)
        del df2, buf
        gc.collect()
    except Exception as e:
        set_progress(0, "Idle")
        return render_template("index.html", upload_error=f"Could not build charts: {e}")

    try:
        os.unlink(tmp_path)
        session.pop("tmp_path", None)
    except:
        pass

    set_progress(90, "Building Dashboard...")

    hs = stats["health_score"]
    health_label = ("Excellent" if hs >= 85 else "Good" if hs >= 70 else "Fair" if hs >= 50 else "Poor")
    health_color = ("green"     if hs >= 85 else "blue" if hs >= 70 else "gold" if hs >= 50 else "red")

    es             = stats["error_summary"]
    critical_count = sum(v["count"] for v in es.values() if v["severity"] == "critical")
    warning_count  = sum(v["count"] for v in es.values() if v["severity"] == "warning")
    total_errors   = sum(v["count"] for k, v in es.items() if not k.startswith("0 -"))

    summary_points = []
    if critical_count > 0:
        summary_points.append({"type": "critical", "icon": "🔴",
            "text": f"{critical_count} critical fault event(s) detected — immediate inspection recommended."})
    if warning_count > 0:
        summary_points.append({"type": "warning", "icon": "🟡",
            "text": f"{warning_count} warning event(s) logged (voltage / pressure / flow limits)."})
    if total_errors == 0:
        summary_points.append({"type": "ok", "icon": "✅",
            "text": "Zero fault conditions recorded in this report period."})
    if hs >= 85:
        summary_points.append({"type": "ok", "icon": "🟢",
            "text": f"Device operating in excellent condition. Health score: {hs}."})
    npct = stats["network_pct_4g"]
    if npct > 80:
        summary_points.append({"type": "ok", "icon": "📶",
            "text": f"Strong 4G connectivity — {npct}% of the time on 4G."})
    elif npct < 40:
        summary_points.append({"type": "warning", "icon": "📶",
            "text": f"Weak 4G connectivity — only {npct}% on 4G. Check network coverage."})
    if stats["start_count"] > 50:
        summary_points.append({"type": "warning", "icon": "⚡",
            "text": f"High motor start count ({stats['start_count']}). Consider investigating frequent cycling."})

    qf = stats["quality_findings"]
    stats["pack_count_total"] = int(stats["pack_count_total"])
    stats["running_time"]     = minutes_to_dhm(stats["motor_running_mins"])
    stats["total_log_time"]   = minutes_to_dhm(stats["total_log_mins"])
    stats["offline_time"]     = minutes_to_dhm(stats["offline_mins"])
    stats["idle_time"]        = minutes_to_dhm(stats["motor_idle_mins"])
    #stats["error_time"]       = minutes_to_dhm(stats["error_running_mins"])
    stats["error_time"]       = minutes_to_dhm(sum(stats["error_duration_by_code"].values()))
    stats["no_error_run_time"]= minutes_to_dhm(stats["no_error_running_mins"])

    set_progress(100, "Dashboard Ready!")

    return render_template(
        "dashboard.html",
        filename=filename,
        filter_applied=use_filter,
        filter_from=date_from if use_filter else min_date,
        filter_to=date_to if use_filter else max_date,
        **stats,
        **charts,
        health_label=health_label, health_color=health_color,
        critical_count=critical_count, warning_count=warning_count, total_errors=total_errors,
        summary_points=summary_points,
        quality_ok=sum(1 for f in qf if f["status"] == "ok"),
        quality_warn=sum(1 for f in qf if f["status"] == "warn"),
        quality_miss=sum(1 for f in qf if f["status"] == "missing"),
    )


if __name__ == "__main__":
    app.run(debug=True)
