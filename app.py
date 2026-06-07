from flask import Flask, render_template, request
import pandas as pd
import os

app = Flask(__name__)

# ═══════════════════════════════════════════════════════
# CONSTANTS
# ═══════════════════════════════════════════════════════

ERROR_CODES = {
    0:      "No Error",
    1:      "Tank Full / Sump Dry",         # Bit 0
    2:      "Low Voltage",                   # Bit 1
    4:      "High Voltage",                  # Bit 2
    8:      "Short Circuit",                 # Bit 3
    16:     "Overload",                      # Bit 4
    32:     "Dry Run",                       # Bit 5
    64:     "Pressure Cut In",               # Bit 6
    128:    "Pressure Cut Off",              # Bit 7
    256:    "Flow Min",                      # Bit 8
    512:    "Flow Max",                      # Bit 9
    1024:   "Pressure Sensor Not Connected", # Bit 10
    2048:   "Flow Sensor Not Connected",     # Bit 11
    4096:   "Phase Sequence Fail",           # Bit 12
    8192:   "Y Phase Fail",                  # Bit 13
    16384:  "Power Failure",                 # Bit 14
    32768:  "Phase Unbalance",               # Bit 15
    65536:  "Current Unbalance",             # Bit 16
    131072: "Pressure Sensor Fault",         # Bit 17
    262144: "R Phase Fail",                  # Bit 18
    524288: "B Phase Fail",                  # Bit 19
    786432: "R/B Phase Fail",                # Bit 20
}

MODE_DESCRIPTIONS = {
    0: "Manual Mode",
    1: "Auto Mode",
    2: "Timer Mode",
    3: "Schedule Mode",
    4: "Bypass Mode",
}

MODE_ICONS = {
    0: "🔧",
    1: "🤖",
    2: "⏱",
    3: "📅",
    4: "⚡",
}

CRITICAL_ERRORS = {8, 16, 32, 4096, 8192, 16384, 32768, 65536, 262144, 524288, 786432}
WARNING_ERRORS  = {1, 2, 4, 64, 128, 256, 512, 1024, 2048, 131072}

MAX_FILE_SIZE_MB   = 20
ALLOWED_EXTENSIONS = {".xlsx", ".xls"}

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER


# ═══════════════════════════════════════════════════════
# HELPER FUNCTIONS
# ═══════════════════════════════════════════════════════

def safe_col(df, col):
    """Return column as numeric list, filling NaN with 0. Empty list if missing."""
    if col not in df.columns:
        return []
    return pd.to_numeric(df[col], errors="coerce").fillna(0).tolist()


def downsample(lst, max_points=500):
    """Reduce list to max_points evenly spaced — keeps charts fast & RAM low."""
    if len(lst) <= max_points:
        return lst
    step = len(lst) // max_points
    return lst[::step][:max_points]


def col_stats(df, col):
    """Return (min, max, avg) rounded to 2 dp. Returns (0,0,0) if missing."""
    if col not in df.columns:
        return 0, 0, 0
    s = pd.to_numeric(df[col], errors="coerce").dropna()
    if s.empty:
        return 0, 0, 0
    return round(float(s.min()), 2), round(float(s.max()), 2), round(float(s.mean()), 2)


def minutes_to_dhm(total_minutes):
    """Convert total minutes → 'XD : HH : MM' string."""
    total_minutes = int(total_minutes)
    d = total_minutes // (24 * 60)
    r = total_minutes % (24 * 60)
    h = r // 60
    m = r % 60
    return f"{d}D : {h:02d}H : {m:02d}M"


def validate_upload(file):
    """Return (ok: bool, message: str)."""
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


def compute_health_score(df, error_summary, avg_voltage, avg_current):
    """
    0–100 health score.
      Error severity   : up to -40 pts
      Voltage balance  : up to -20 pts
      Current balance  : up to -20 pts
      Data completeness: up to -20 pts
    """
    score = 100.0

    # Error deductions
    for key, val in error_summary.items():
        code_str = key.split(" - ")[0]
        try:
            code = int(code_str)
        except ValueError:
            continue
        if code == 0:
            continue
        count = val["count"]
        if code in CRITICAL_ERRORS:
            score -= min(15, count * 3)
        elif code in WARNING_ERRORS:
            score -= min(8,  count * 1)
        else:
            score -= min(4,  count * 0.5)

    # Voltage balance (ideal ~415 V three-phase)
    if avg_voltage > 0:
        deviation = abs(avg_voltage - 415) / 415
        score -= min(20, deviation * 60)

    # Current balance
    if all(c in df.columns for c in ["Current Amp", "Current Amp2", "Current Amp3"]):
        means   = [pd.to_numeric(df["Current Amp"],  errors="coerce").mean(),
                   pd.to_numeric(df["Current Amp2"], errors="coerce").mean(),
                   pd.to_numeric(df["Current Amp3"], errors="coerce").mean()]
        overall = sum(means) / 3
        if overall > 0:
            imbalance = max(abs(m - overall) / overall for m in means)
            score -= min(20, imbalance * 40)

    # Data completeness
    key_cols = [
        "Line Voltage", "Line Voltage 2", "Line Voltage 3",
        "Current Amp",  "Current Amp2",   "Current Amp3",
        "Pressure",     "Flow Sensor",    "Frequency",
    ]
    present = sum(1 for c in key_cols if c in df.columns)
    score -= (1 - present / len(key_cols)) * 20

    return max(0, min(100, round(score, 1)))


def data_quality_check(df):
    """Return list of quality findings for key sensor columns."""
    findings = []
    total    = len(df)

    def pct(n):
        return round(n / total * 100, 1) if total else 0

    checks = [
        ("Line Voltage",   "R-Phase Voltage"),
        ("Line Voltage 2", "Y-Phase Voltage"),
        ("Line Voltage 3", "B-Phase Voltage"),
        ("Current Amp",    "R-Phase Current"),
        ("Current Amp2",   "Y-Phase Current"),
        ("Current Amp3",   "B-Phase Current"),
        ("Pressure",       "Pressure"),
        ("Flow Sensor",    "Flow Sensor"),
        ("Frequency",      "Frequency"),
        ("Signal",         "Signal Strength"),
    ]

    for col, label in checks:
        if col not in df.columns:
            findings.append({"field": label, "status": "missing",
                             "message": "Column not found in dataset", "pct": 0})
            continue
        null_pct = pct(int(df[col].isna().sum()))
        zero_pct = pct(int((df[col].fillna(0) == 0).sum()))
        if null_pct > 10:
            findings.append({"field": label, "status": "warn",
                             "message": f"{null_pct}% missing values", "pct": 100 - null_pct})
        elif zero_pct > 50:
            findings.append({"field": label, "status": "warn",
                             "message": f"{zero_pct}% zero values", "pct": 100 - zero_pct})
        else:
            findings.append({"field": label, "status": "ok",
                             "message": f"Good — {100 - null_pct}% complete", "pct": 100 - null_pct})

    return findings


# ═══════════════════════════════════════════════════════
# ROUTES
# ═══════════════════════════════════════════════════════

@app.route("/")
def home():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():

    file = request.files.get("file")

    # ── Secure upload validation ───────────────────────
    ok, msg = validate_upload(file)
    if not ok:
        return render_template("index.html", upload_error=msg)

    path = os.path.join(app.config["UPLOAD_FOLDER"], file.filename)
    file.save(path)

    try:
        ext    = os.path.splitext(file.filename)[1].lower()
        engine = "xlrd" if ext == ".xls" else "openpyxl"

        # Read only the columns we actually need — saves RAM on Render free plan
        NEEDED_COLS = [
            "DeviceId", "IoTHubName", "PumpPhaseType", "QueuedTime-IST",
            "MotorRunningStatus", "Error CondMon", "ModeOfOperating",
            "Line Voltage", "Line Voltage 2", "Line Voltage 3",
            "Current Amp", "Current Amp2", "Current Amp3",
            "Pressure", "Flow Sensor", "Frequency", "Signal",
            "PackCount", "NetType", "Total Running Time",
        ]

        # First pass — find which needed columns exist
        header_df = pd.read_excel(path, engine=engine, nrows=0)
        use_cols  = [c for c in NEEDED_COLS if c in header_df.columns]

        # Second pass — read only those columns, dtype=str avoids datetime crash
        df = pd.read_excel(path, engine=engine, usecols=use_cols, dtype=str)

    except Exception as e:
        return render_template("index.html", upload_error=f"Could not read file: {e}")

    total_rows = len(df)

    # ── Device information ─────────────────────────────
    device_id  = str(df["DeviceId"].iloc[0])     if "DeviceId"      in df.columns else "—"
    iot_hub    = str(df["IoTHubName"].iloc[0])    if "IoTHubName"    in df.columns else "—"
    _pump_phase_raw = str(df["PumpPhaseType"].iloc[0]) if "PumpPhaseType" in df.columns else "—"
    pump_phase = (
        "Three Phase Pump" if _pump_phase_raw == "1" else
        "Single Phase Pump" if _pump_phase_raw == "0" else
        _pump_phase_raw
    )

    date_range = "—"
    if "QueuedTime-IST" in df.columns:
        ts = pd.to_datetime(df["QueuedTime-IST"], errors="coerce").dropna()
        if not ts.empty:
            date_range = f"{ts.min().strftime('%d %b %Y')} → {ts.max().strftime('%d %b %Y')}"

    # ── Motor start count ─────────────────────────────
    start_count = 0
    if "MotorRunningStatus" in df.columns:
        # Handle both string ("True"/"False"/"1"/"0") and numeric (0/1)
        raw = df["MotorRunningStatus"].fillna(0).astype(str).str.strip().str.upper()
        ms  = raw.map(lambda x: x in ("1", "TRUE", "YES", "ON"))
        start_count = int(((ms.shift(1) == False) & (ms == True)).sum())

    # ── Error summary ──────────────────────────────────
    error_summary = {}
    if "Error CondMon" in df.columns:
        for code, count in df["Error CondMon"].value_counts().to_dict().items():
            code     = int(code)
            desc     = ERROR_CODES.get(code, "Unknown Error")
            severity = (
                "critical" if code in CRITICAL_ERRORS else
                "warning"  if code in WARNING_ERRORS  else
                "info"
            )
            error_summary[f"{code} - {desc}"] = {"count": count, "severity": severity}

    # ── Mode summary ───────────────────────────────────
    mode_summary = {}
    if "ModeOfOperating" in df.columns:
        for code, count in df["ModeOfOperating"].value_counts().to_dict().items():
            try:
                code_int = int(code)
            except (ValueError, TypeError):
                code_int = -1
            desc = MODE_DESCRIPTIONS.get(code_int, f"Unknown Mode ({code})")
            icon = MODE_ICONS.get(code_int, "❓")
            mode_summary[desc] = {"code": code_int, "count": count, "icon": icon}

    # ── Time labels ────────────────────────────────────
    time_labels = []
    if "QueuedTime-IST" in df.columns:
        time_labels = df["QueuedTime-IST"].astype(str).tolist()

    # ── Voltage ────────────────────────────────────────
    voltage1 = safe_col(df, "Line Voltage")
    voltage2 = safe_col(df, "Line Voltage 2")
    voltage3 = safe_col(df, "Line Voltage 3")

    v1_min, v1_max, _ = col_stats(df, "Line Voltage")
    v2_min, v2_max, _ = col_stats(df, "Line Voltage 2")
    v3_min, v3_max, _ = col_stats(df, "Line Voltage 3")

    avg_voltage = 0
    if all(c in df.columns for c in ["Line Voltage", "Line Voltage 2", "Line Voltage 3"]):
        avg_voltage = round(
            (pd.to_numeric(df["Line Voltage"], errors="coerce").mean() +
             pd.to_numeric(df["Line Voltage 2"], errors="coerce").mean() +
             pd.to_numeric(df["Line Voltage 3"], errors="coerce").mean()) / 3, 2
        )

    # ── Current ────────────────────────────────────────
    current1 = safe_col(df, "Current Amp")
    current2 = safe_col(df, "Current Amp2")
    current3 = safe_col(df, "Current Amp3")

    c1_min, c1_max, _ = col_stats(df, "Current Amp")
    c2_min, c2_max, _ = col_stats(df, "Current Amp2")
    c3_min, c3_max, _ = col_stats(df, "Current Amp3")

    avg_current = 0
    if all(c in df.columns for c in ["Current Amp", "Current Amp2", "Current Amp3"]):
        avg_current = round(
            (pd.to_numeric(df["Current Amp"],  errors="coerce").mean() +
             pd.to_numeric(df["Current Amp2"], errors="coerce").mean() +
             pd.to_numeric(df["Current Amp3"], errors="coerce").mean()) / 3, 2
        )

    # ── Process parameters ─────────────────────────────
    pressure  = safe_col(df, "Pressure")
    flow      = safe_col(df, "Flow Sensor")
    frequency = safe_col(df, "Frequency")
    signal    = safe_col(df, "Signal")

    p_min,    p_max,    p_avg    = col_stats(df, "Pressure")
    f_min,    f_max,    f_avg    = col_stats(df, "Flow Sensor")
    freq_min, freq_max, freq_avg = col_stats(df, "Frequency")

    # ── Pack count (vectorized) ────────────────────────
    pack_count_total = 0
    if "PackCount" in df.columns:
        pack_data = pd.to_numeric(df["PackCount"], errors="coerce").fillna(0)
        diff = pack_data.diff().fillna(0)
        pack_count_total = int(diff[diff > 0].sum())

    pack_count_graph = safe_col(df, "PackCount")

    # ── Network type ───────────────────────────────────
    network_numeric = []
    net_4g_count    = 0
    net_2g_count    = 0
    if "NetType" in df.columns:
        for n in df["NetType"]:
            if str(n).upper() == "4G":
                network_numeric.append(4)
                net_4g_count += 1
            else:
                network_numeric.append(2)
                net_2g_count += 1

    network_pct_4g = round(net_4g_count / max(net_4g_count + net_2g_count, 1) * 100, 1)

    # ── Total running time (vectorized) ───────────────
    total_runtime = 0
    if "Total Running Time" in df.columns:
        rt   = pd.to_numeric(df["Total Running Time"], errors="coerce").fillna(0)
        diff = rt.diff().fillna(0)
        total_runtime = float(diff[diff > 0].sum())

    running_time = minutes_to_dhm(total_runtime)

    # ── Health score ───────────────────────────────────
    health_score = compute_health_score(df, error_summary, avg_voltage, avg_current)
    health_label = (
        "Excellent" if health_score >= 85 else
        "Good"      if health_score >= 70 else
        "Fair"      if health_score >= 50 else
        "Poor"
    )
    health_color = (
        "green" if health_score >= 85 else
        "blue"  if health_score >= 70 else
        "gold"  if health_score >= 50 else
        "red"
    )

    # ── Data quality check ─────────────────────────────
    quality_findings = data_quality_check(df)
    quality_ok   = sum(1 for f in quality_findings if f["status"] == "ok")
    quality_warn = sum(1 for f in quality_findings if f["status"] == "warn")
    quality_miss = sum(1 for f in quality_findings if f["status"] == "missing")

    # ── Executive summary ──────────────────────────────
    critical_count = sum(v["count"] for v in error_summary.values() if v["severity"] == "critical")
    warning_count  = sum(v["count"] for v in error_summary.values() if v["severity"] == "warning")
    total_errors   = sum(v["count"] for k, v in error_summary.items() if not k.startswith("0 -"))

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
    if health_score >= 85:
        summary_points.append({"type": "ok", "icon": "🟢",
            "text": f"Device operating in excellent condition. Health score: {health_score}."})
    if network_pct_4g > 80:
        summary_points.append({"type": "ok", "icon": "📶",
            "text": f"Strong 4G connectivity — {network_pct_4g}% of the time on 4G."})
    elif network_pct_4g < 40:
        summary_points.append({"type": "warning", "icon": "📶",
            "text": f"Weak 4G connectivity — only {network_pct_4g}% on 4G. Check network coverage."})
    if start_count > 50:
        summary_points.append({"type": "warning", "icon": "⚡",
            "text": f"High motor start count ({start_count}). Consider investigating frequent cycling."})

    # ── Downsample chart data (max 500 pts) — saves RAM & response size ──
    ds_labels          = downsample(time_labels)
    ds_voltage1        = downsample(voltage1)
    ds_voltage2        = downsample(voltage2)
    ds_voltage3        = downsample(voltage3)
    ds_current1        = downsample(current1)
    ds_current2        = downsample(current2)
    ds_current3        = downsample(current3)
    ds_pressure        = downsample(pressure)
    ds_flow            = downsample(flow)
    ds_frequency       = downsample(frequency)
    ds_signal          = downsample(signal)
    ds_network_numeric = downsample(network_numeric)
    ds_pack_graph      = downsample(pack_count_graph)

    # ── Render dashboard ───────────────────────────────
    return render_template(
        "dashboard.html",

        # Device info
        device_id=device_id,
        iot_hub=iot_hub,
        pump_phase=pump_phase,
        date_range=date_range,
        total_rows=total_rows,
        filename=file.filename,

        # KPIs
        start_count=start_count,
        running_time=running_time,
        avg_voltage=avg_voltage,
        avg_current=avg_current,
        pack_count_total=int(pack_count_total),

        # Voltage min/max
        v1_min=v1_min, v1_max=v1_max,
        v2_min=v2_min, v2_max=v2_max,
        v3_min=v3_min, v3_max=v3_max,

        # Current min/max
        c1_min=c1_min, c1_max=c1_max,
        c2_min=c2_min, c2_max=c2_max,
        c3_min=c3_min, c3_max=c3_max,

        # Process stats
        p_min=p_min,   p_max=p_max,   p_avg=p_avg,
        f_min=f_min,   f_max=f_max,   f_avg=f_avg,
        freq_min=freq_min, freq_max=freq_max, freq_avg=freq_avg,

        # Health
        health_score=health_score,
        health_label=health_label,
        health_color=health_color,

        # Executive summary
        summary_points=summary_points,
        critical_count=critical_count,
        warning_count=warning_count,
        total_errors=total_errors,

        # Tables
        error_summary=error_summary,
        mode_summary=mode_summary,

        # Data quality
        quality_findings=quality_findings,
        quality_ok=quality_ok,
        quality_warn=quality_warn,
        quality_miss=quality_miss,

        # Network
        net_4g_count=net_4g_count,
        net_2g_count=net_2g_count,
        network_pct_4g=network_pct_4g,

        # Chart data (downsampled)
        labels=ds_labels,
        voltage1=ds_voltage1, voltage2=ds_voltage2, voltage3=ds_voltage3,
        current1=ds_current1, current2=ds_current2, current3=ds_current3,
        pressure=ds_pressure, flow=ds_flow,
        frequency=ds_frequency, signal=ds_signal,
        network_numeric=ds_network_numeric,
        pack_count_graph=ds_pack_graph,
    )


if __name__ == "__main__":
    app.run(debug=True)
