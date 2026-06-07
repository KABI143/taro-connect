from flask import Flask, render_template, request
import pandas as pd
import io
import os
import gc

app = Flask(__name__)

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

CRITICAL_ERRORS = {8, 16, 32, 4096, 8192, 16384, 32768, 65536, 262144, 524288, 786432}
WARNING_ERRORS  = {1, 2, 4, 64, 128, 256, 512, 1024, 2048, 131072}

MAX_FILE_SIZE_MB   = 20
ALLOWED_EXTENSIONS = {".xlsx", ".xls"}

VOLTAGE_COLS    = ["Line Voltage", "Line Voltage 2", "Line Voltage 3"]
CURRENT_COLS    = ["Current Amp", "Current Amp2", "Current Amp3"]
KEY_SENSOR_COLS = VOLTAGE_COLS + CURRENT_COLS + ["Pressure", "Flow Sensor", "Frequency"]

# Pass 1 — stats only (scalars needed for health score, KPIs, summaries)
STATS_COLS = [
    "DeviceId", "IoTHubName", "PumpPhaseType", "QueuedTime-IST",
    "MotorRunningStatus", "Error CondMon", "ModeOfOperating",
    "Total Running Time", "PackCount", "NetType",
] + KEY_SENSOR_COLS + ["Signal"]

# Pass 2 — chart data only (downsampled to 500 pts)
CHART_COLS = ["QueuedTime-IST"] + KEY_SENSOR_COLS + ["Signal", "PackCount", "NetType"]

CHART_DOWNSAMPLE = 500

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
    """Read only the columns that exist in the file."""
    buf.seek(0)
    header = pd.read_excel(buf, engine=engine, nrows=0)
    use    = [c for c in wanted_cols if c in header.columns]
    del header
    buf.seek(0)
    return pd.read_excel(buf, engine=engine, usecols=use)


def col_stats(series):
    """(min, max, avg) from a Series. (0,0,0) if empty."""
    s = pd.to_numeric(series, errors="coerce").dropna()
    if s.empty:
        return 0, 0, 0
    return round(float(s.min()), 2), round(float(s.max()), 2), round(float(s.mean()), 2)


def group_avg(df, cols):
    present = [c for c in cols if c in df.columns]
    if not present:
        return 0
    return round(sum(pd.to_numeric(df[c], errors="coerce").mean() for c in present) / len(present), 2)


def minutes_to_dhm(total_minutes):
    total_minutes = int(total_minutes)
    d = total_minutes // 1440
    h = (total_minutes % 1440) // 60
    m = total_minutes % 60
    return f"{d}D : {h:02d}H : {m:02d}M"


def downsample_series(series, max_pts=CHART_DOWNSAMPLE):
    """Downsample a Series to list — no full intermediate list."""
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
# PASS 1 — compute all scalars from stats columns
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

    start_count = 0
    if "MotorRunningStatus" in df.columns:
        raw = df["MotorRunningStatus"].fillna(0).astype(str).str.strip().str.upper()
        ms  = raw.isin({"1", "TRUE", "YES", "ON"})
        start_count = int((~ms.shift(1, fill_value=False) & ms).sum())

    error_summary = {}
    if "Error CondMon" in df.columns:
        codes = pd.to_numeric(df["Error CondMon"], errors="coerce").dropna().astype(int)
        for code, count in codes.value_counts().items():
            desc     = ERROR_CODES.get(code, "Unknown Error")
            severity = ("critical" if code in CRITICAL_ERRORS else
                        "warning"  if code in WARNING_ERRORS  else "info")
            error_summary[f"{code} - {desc}"] = {"count": count, "severity": severity}

    mode_summary = {}
    if "ModeOfOperating" in df.columns:
        for code, count in df["ModeOfOperating"].value_counts().items():
            try:    code_int = int(code)
            except: code_int = -1
            mode_summary[MODE_DESCRIPTIONS.get(code_int, f"Unknown Mode ({code})")] = {
                "code": code_int, "count": count, "icon": MODE_ICONS.get(code_int, "❓")
            }

    def _s(col): return df[col] if col in df.columns else pd.Series(dtype=float)

    v1_min, v1_max, _ = col_stats(_s("Line Voltage"))
    v2_min, v2_max, _ = col_stats(_s("Line Voltage 2"))
    v3_min, v3_max, _ = col_stats(_s("Line Voltage 3"))
    avg_voltage        = group_avg(df, VOLTAGE_COLS)

    c1_min, c1_max, _ = col_stats(_s("Current Amp"))
    c2_min, c2_max, _ = col_stats(_s("Current Amp2"))
    c3_min, c3_max, _ = col_stats(_s("Current Amp3"))
    avg_current        = group_avg(df, CURRENT_COLS)

    p_min,    p_max,    p_avg    = col_stats(_s("Pressure"))
    f_min,    f_max,    f_avg    = col_stats(_s("Flow Sensor"))
    freq_min, freq_max, freq_avg = col_stats(_s("Frequency"))

    pack_count_total = 0
    if "PackCount" in df.columns:
        diff = pd.to_numeric(df["PackCount"], errors="coerce").diff().fillna(0)
        pack_count_total = int(diff[diff > 0].sum())

    net_4g_count = net_2g_count = 0
    if "NetType" in df.columns:
        is_4g        = df["NetType"].astype(str).str.upper() == "4G"
        net_4g_count = int(is_4g.sum())
        net_2g_count = len(df) - net_4g_count
    network_pct_4g = round(net_4g_count / max(net_4g_count + net_2g_count, 1) * 100, 1)

    total_runtime = 0
    if "Total Running Time" in df.columns:
        diff = pd.to_numeric(df["Total Running Time"], errors="coerce").diff().fillna(0)
        total_runtime = float(diff[diff > 0].sum())

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
        v1_min=v1_min, v1_max=v1_max, v2_min=v2_min, v2_max=v2_max,
        v3_min=v3_min, v3_max=v3_max, avg_voltage=avg_voltage,
        c1_min=c1_min, c1_max=c1_max, c2_min=c2_min, c2_max=c2_max,
        c3_min=c3_min, c3_max=c3_max, avg_current=avg_current,
        p_min=p_min, p_max=p_max, p_avg=p_avg,
        f_min=f_min, f_max=f_max, f_avg=f_avg,
        freq_min=freq_min, freq_max=freq_max, freq_avg=freq_avg,
        pack_count_total=pack_count_total, net_4g_count=net_4g_count,
        net_2g_count=net_2g_count, network_pct_4g=network_pct_4g,
        total_runtime=total_runtime, health_score=health_score,
        quality_findings=quality_findings,
    )


# ═══════════════════════════════════════════════════════
# PASS 2 — chart data only (downsampled)
# ═══════════════════════════════════════════════════════

def compute_charts(df):
    def _ds(col):
        if col not in df.columns:
            return []
        return downsample_series(df[col])

    time_labels = []
    if "QueuedTime-IST" in df.columns:
        time_labels = downsample_list(df["QueuedTime-IST"].astype(str).tolist())

    net_numeric = []
    if "NetType" in df.columns:
        net_numeric = downsample_list(
            (df["NetType"].astype(str).str.upper() == "4G").map({True: 4, False: 2}).tolist()
        )

    return dict(
        labels=time_labels,
        voltage1=_ds("Line Voltage"),   voltage2=_ds("Line Voltage 2"), voltage3=_ds("Line Voltage 3"),
        current1=_ds("Current Amp"),    current2=_ds("Current Amp2"),   current3=_ds("Current Amp3"),
        pressure=_ds("Pressure"),       flow=_ds("Flow Sensor"),
        frequency=_ds("Frequency"),     signal=_ds("Signal"),
        pack_count_graph=_ds("PackCount"),
        network_numeric=net_numeric,
    )


# ═══════════════════════════════════════════════════════
# ROUTES
# ═══════════════════════════════════════════════════════

@app.route("/")
def home():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():

    file = request.files.get("file")
    ok, msg = validate_upload(file)
    if not ok:
        return render_template("index.html", upload_error=msg)

    filename = file.filename
    try:
        ext    = os.path.splitext(filename)[1].lower()
        engine = "xlrd" if ext == ".xls" else "openpyxl"
        buf    = io.BytesIO(file.read())
        del file
    except Exception as e:
        return render_template("index.html", upload_error=f"Could not read file: {e}")

    # ── PASS 1 : stats columns only ───────────────────
    try:
        df1   = read_excel_cols(buf, engine, STATS_COLS)
        stats = compute_stats(df1)
        del df1
        gc.collect()
    except Exception as e:
        return render_template("index.html", upload_error=f"Could not process file: {e}")

    # ── PASS 2 : chart columns only ───────────────────
    try:
        df2    = read_excel_cols(buf, engine, CHART_COLS)
        charts = compute_charts(df2)
        del df2, buf
        gc.collect()
    except Exception as e:
        return render_template("index.html", upload_error=f"Could not build charts: {e}")

    # ── Derived display values ─────────────────────────
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

    # Override computed fields before spreading stats
    stats["pack_count_total"] = int(stats["pack_count_total"])
    stats["running_time"]     = minutes_to_dhm(stats["total_runtime"])

    return render_template(
        "dashboard.html",
        filename=filename,
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
