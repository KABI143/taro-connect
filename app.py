"""
TARO CONNECT - PUMP ANALYTICS DASHBOARD
Production-Grade Flask Application with Comprehensive Error Handling
Version 2.2 - Bug Fixed
"""

from flask import Flask, render_template, request, session, jsonify
import pandas as pd
import numpy as np
import io
import os
import gc
import tempfile
import logging
import threading
from datetime import datetime
from functools import wraps

app = Flask(__name__)
app.secret_key = "taroconnect-secret-key-2024-production"
app.config['MAX_CONTENT_LENGTH'] = 25 * 1024 * 1024

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('taro_connect.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

_progress_lock = threading.Lock()
progress_data = {"percent": 0, "status": "Idle"}

def set_progress(percent, status):
    try:
        with _progress_lock:
            progress_data["percent"] = min(100, max(0, percent))
            progress_data["status"] = str(status)
            logger.info(f"Progress: {percent}% - {status}")
    except Exception as e:
        logger.warning(f"Progress update failed: {e}")

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

MODE_DESCRIPTIONS = {
    0: "Manual Mode",
    1: "Auto Mode",
    2: "Timer Mode",
    3: "Schedule Mode",
    4: "Bypass Mode"
}
MODE_ICONS = {0: "🔧", 1: "🤖", 2: "⏱", 3: "📅", 4: "⚡"}

CRITICAL_ERRORS = {8, 16}
WARNING_ERRORS = {
    1, 2, 4, 32, 64, 128, 256, 512, 1024, 2048, 4096, 8192,
    16384, 32768, 65536, 131072, 262144, 524288, 786432
}

MAX_FILE_SIZE_MB = 25
ALLOWED_EXTENSIONS = {".xlsx", ".xls"}

VOLTAGE_COLS = ["Line Voltage", "Line Voltage 2", "Line Voltage 3"]
CURRENT_COLS = ["Current Amp", "Current Amp2", "Current Amp3"]
KEY_SENSOR_COLS = VOLTAGE_COLS + CURRENT_COLS + ["Pressure", "Flow Sensor", "Frequency"]

OFFLINE_GAP_MINUTES = 3
DOWNSAMPLE_MAX_POINTS = 999999

STATS_COLS = [
    "DeviceId", "IoTHubName", "PumpPhaseType", "QueuedTime-IST",
    "MotorRunningStatus", "Error CondMon", "ModeOfOperating",
    "Total Running Time", "PackCount", "NetType",
] + KEY_SENSOR_COLS + ["Signal"]

CHART_COLS = ["QueuedTime-IST"] + KEY_SENSOR_COLS + [
    "Signal", "PackCount", "NetType", "MotorRunningStatus", "Error CondMon"
]


# ═══════════════════════════════════════════════════════
# BUG FIX 1: parse_ist_column — handle already-parsed datetime columns
# ═══════════════════════════════════════════════════════
def parse_ist_column(series):
    """
    Parse IST timestamp column with fallback handling.
    FIX: Check if already datetime64 first to avoid mis-parsing.
    """
    if series.empty:
        return series

    # FIX: If already a datetime column (e.g. read directly by pandas),
    # return as-is instead of trying float/Excel-serial conversion.
    if pd.api.types.is_datetime64_any_dtype(series):
        return series

    try:
        sample = series.dropna().iloc[0] if len(series.dropna()) > 0 else None
        if sample is None:
            return pd.to_datetime(series, errors='coerce')

        try:
            float(sample)
            # Only treat as Excel date serial if it's actually a bare number
            return pd.to_datetime(series, unit='D', origin='1899-12-30', errors='coerce')
        except (ValueError, TypeError):
            return pd.to_datetime(series, errors='coerce')
    except Exception as e:
        logger.warning(f"Timestamp parsing failed: {e}")
        return pd.to_datetime(series, errors='coerce')


def validate_upload(file):
    try:
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
    except Exception as e:
        logger.error(f"File validation error: {e}")
        return False, f"Validation error: {str(e)}"


def read_excel_cols(buf, engine, wanted_cols):
    try:
        buf.seek(0)
        header = pd.read_excel(buf, engine=engine, nrows=0)
        use = [c for c in wanted_cols if c in header.columns]
        del header
        gc.collect()
        buf.seek(0)
        return pd.read_excel(buf, engine=engine, usecols=use)
    except Exception as e:
        logger.error(f"Excel read error: {e}")
        raise


def col_stats(series):
    try:
        s = pd.to_numeric(series, errors="coerce").dropna()
        if s.empty:
            return 0.0, 0.0, 0.0
        return (
            round(float(s.min()), 2),
            round(float(s.max()), 2),
            round(float(s.mean()), 2)
        )
    except Exception as e:
        logger.warning(f"Column stats calculation failed: {e}")
        return 0.0, 0.0, 0.0


def col_stats_motor_on(df, col):
    try:
        if col not in df.columns or "MotorRunningStatus" not in df.columns:
            return 0.0, 0.0, 0.0
        motor_on = (df["MotorRunningStatus"]
                    .astype(str).str.strip().str.upper()
                    .isin(["1", "TRUE", "YES", "ON"]))
        s = pd.to_numeric(df.loc[motor_on, col], errors="coerce").dropna()
        if s.empty:
            return 0.0, 0.0, 0.0
        return (
            round(float(s.min()), 2),
            round(float(s.max()), 2),
            round(float(s.mean()), 2)
        )
    except Exception as e:
        logger.warning(f"Motor-on stats failed for {col}: {e}")
        return 0.0, 0.0, 0.0


def group_avg(df, cols):
    try:
        present = [c for c in cols if c in df.columns]
        if not present:
            return 0.0
        vals = [pd.to_numeric(df[c], errors="coerce").mean() for c in present]
        vals = [v for v in vals if not pd.isna(v)]
        if not vals:
            return 0.0
        return round(sum(vals) / len(vals), 2)
    except Exception as e:
        logger.warning(f"Group average calculation failed: {e}")
        return 0.0


def group_avg_motor_on(df, cols):
    try:
        if "MotorRunningStatus" not in df.columns:
            return group_avg(df, cols)
        motor_on = (df["MotorRunningStatus"]
                    .astype(str).str.strip().str.upper()
                    .isin(["1", "TRUE", "YES", "ON"]))
        present = [c for c in cols if c in df.columns]
        if not present:
            return 0.0
        vals = [pd.to_numeric(df.loc[motor_on, c], errors="coerce").mean() for c in present]
        vals = [v for v in vals if not pd.isna(v)]
        if not vals:
            return 0.0
        return round(float(sum(vals) / len(vals)), 2)
    except Exception as e:
        logger.warning(f"Motor-on group average failed: {e}")
        return 0.0


def calc_avg_voltage_no_fail(df, cols):
    try:
        phase_fail_codes = {4096, 8192, 262144, 524288, 786432}
        power_fail_code = 16384
        fail_codes = phase_fail_codes | {power_fail_code}
        if "Error CondMon" not in df.columns:
            return group_avg(df, cols)
        error_col = (pd.to_numeric(df["Error CondMon"], errors="coerce")
                     .fillna(0).astype(int))
        no_fail_mask = ~error_col.isin(fail_codes)
        present = [c for c in cols if c in df.columns]
        if not present:
            return 0.0
        vals = [pd.to_numeric(df.loc[no_fail_mask, c], errors="coerce").mean() for c in present]
        vals = [v for v in vals if not pd.isna(v)]
        if not vals:
            return 0.0
        return round(sum(vals) / len(vals), 2)
    except Exception as e:
        logger.warning(f"Voltage no-fail calculation failed: {e}")
        return 0.0


def minutes_to_dhm(total_minutes):
    try:
        total_seconds = int(round(total_minutes * 60))
        days = total_seconds // 86400
        hours = (total_seconds % 86400) // 3600
        mins = (total_seconds % 3600) // 60
        secs = total_seconds % 60
        return f"{days}D : {hours:02d}H : {mins:02d}M : {secs:02d}S"
    except Exception as e:
        logger.warning(f"Time conversion failed: {e}")
        return "0D : 00H : 00M : 00S"


def downsample_series(series, max_pts=DOWNSAMPLE_MAX_POINTS):
    """
    BUG FIX 7: Always convert to plain Python float/int so Flask's JSON
    encoder never sees numpy.bool_, numpy.int64, or Python bool — all of
    which raise 'Object of type bool is not JSON serializable'.
    """
    try:
        n = len(series)
        if n <= max_pts:
            numeric = pd.to_numeric(series, errors="coerce").fillna(0)
        else:
            step = max(1, n // max_pts)
            numeric = pd.to_numeric(series.iloc[::step].iloc[:max_pts], errors="coerce").fillna(0)
        return [float(v) for v in numeric]
    except Exception as e:
        logger.warning(f"Series downsampling failed: {e}")
        return []


def downsample_list(lst, max_pts=DOWNSAMPLE_MAX_POINTS):
    try:
        if len(lst) <= max_pts:
            return lst
        step = max(1, len(lst) // max_pts)
        return lst[::step][:max_pts]
    except Exception as e:
        logger.warning(f"List downsampling failed: {e}")
        return lst


# ═══════════════════════════════════════════════════════
# BUG FIX 2: compute_durations
#   - Motor OFF + Error != 0 time was silently dropped (lost time bug)
#   - total_log_mins was incorrectly including offline_mins
# ═══════════════════════════════════════════════════════
def compute_durations(df):
    result = dict(
        total_log_mins=0.0, motor_running_mins=0.0, motor_idle_mins=0.0,
        no_error_running_mins=0.0, error_running_mins=0.0, offline_mins=0.0,
        error_duration_by_code={}, mode_runtime={},
    )

    needs = {"QueuedTime-IST", "MotorRunningStatus", "Error CondMon", "ModeOfOperating"}
    if not needs.issubset(df.columns):
        logger.warning(f"Missing columns for duration calculation: {needs - set(df.columns)}")
        return result

    try:
        df = df.copy()
        df["QueuedTime-IST"] = parse_ist_column(df["QueuedTime-IST"])
        df = df.sort_values("QueuedTime-IST").reset_index(drop=True)

        ts = df["QueuedTime-IST"]
        motor = (df["MotorRunningStatus"]
                 .astype(str).str.strip().str.upper()
                 .isin(["1", "TRUE", "YES", "ON"]))
        err = pd.to_numeric(df["Error CondMon"], errors="coerce").fillna(0).astype(int)
        mode_s = pd.to_numeric(df["ModeOfOperating"], errors="coerce").fillna(-1).astype(int)

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
            err_code = int(err.iloc[i])
            mode_code = int(mode_s.iloc[i])

            if err_code != 0:
                ec = result["error_duration_by_code"]
                ec[err_code] = ec.get(err_code, 0) + diff

            mr = result["mode_runtime"]
            mr[mode_code] = mr.get(mode_code, 0) + diff

            if is_running:
                result["motor_running_mins"] += diff
                if err_code != 0:
                    result["error_running_mins"] += diff
                else:
                    result["no_error_running_mins"] += diff
            else:
                if err_code == 0:
                    result["motor_idle_mins"] += diff

        logger.info(f"Duration calculation complete: {result['total_log_mins']:.1f} min total")
        return result

    except Exception as e:
        logger.error(f"Duration calculation failed: {e}")
        return result


def compute_error_events(df):
    events = []
    needs = {"QueuedTime-IST", "Error CondMon"}
    if not needs.issubset(df.columns):
        logger.warning("Missing columns for error event detection")
        return events

    try:
        df = df.copy()
        df["QueuedTime-IST"] = parse_ist_column(df["QueuedTime-IST"])
        df = df.sort_values("QueuedTime-IST").reset_index(drop=True)

        err_col = pd.to_numeric(df["Error CondMon"], errors="coerce").fillna(0).astype(int)
        i = 0
        n = len(df)

        while i < n:
            code = int(err_col.iloc[i])
            if code == 0:
                i += 1
                continue
            j = i + 1
            while j < n and int(err_col.iloc[j]) == code:
                j += 1
            t_start = df["QueuedTime-IST"].iloc[i]
            t_end = df["QueuedTime-IST"].iloc[j - 1]
            dur_mins = 0.0
            if pd.notna(t_start) and pd.notna(t_end):
                dur_mins = max(0, (t_end - t_start).total_seconds() / 60)

            def _snap(col):
                if col not in df.columns:
                    return None
                try:
                    val = pd.to_numeric(df[col].iloc[i], errors="coerce")
                    return round(float(val), 2) if not pd.isna(val) else None
                except Exception:
                    return None

            events.append({
                "code": code,
                "desc": ERROR_CODES.get(code, "Unknown Error"),
                "severity": (
                    "critical" if code in CRITICAL_ERRORS
                    else "warning" if code in WARNING_ERRORS
                    else "info"
                ),
                "start": t_start.strftime("%d %b %Y %H:%M:%S") if pd.notna(t_start) else "—",
                "end": t_end.strftime("%d %b %Y %H:%M:%S") if pd.notna(t_end) else "—",
                "duration": minutes_to_dhm(dur_mins),
                "dur_mins": round(dur_mins, 1),
                "v1": _snap("Line Voltage"),
                "v2": _snap("Line Voltage 2"),
                "v3": _snap("Line Voltage 3"),
                "current": _snap("Current Amp"),
            })
            i = j

        sev_order = {"critical": 0, "warning": 1, "info": 2}
        events.sort(key=lambda e: (sev_order.get(e["severity"], 3), e["start"]))
        logger.info(f"Detected {len(events)} error events")
        return events

    except Exception as e:
        logger.error(f"Error event detection failed: {e}")
        return events


def get_device_running_mins(df):
    try:
        if "Total Running Time" not in df.columns:
            return None
        trt = pd.to_numeric(df["Total Running Time"], errors="coerce").dropna()
        if len(trt) < 2:
            return None
        diffs = trt.diff().dropna()
        if (diffs < 0).any():
            positive_increments = diffs[diffs > 0].sum()
            last_val = trt.iloc[-1]
            min_val = trt.min()
            last_segment = last_val - min_val if last_val >= min_val else 0
            return float(positive_increments + last_segment)
        else:
            return float(trt.iloc[-1] - trt.iloc[0])
    except Exception as e:
        logger.warning(f"Running time extraction failed: {e}")
        return None


def compute_stats(df):
    stats = {}

    try:
        total_rows = len(df)
        logger.info(f"Computing stats for {total_rows} rows")
        stats["total_rows"] = total_rows

        stats["device_id"] = str(df["DeviceId"].iloc[0]) if "DeviceId" in df.columns else "—"
        stats["iot_hub"] = str(df["IoTHubName"].iloc[0]) if "IoTHubName" in df.columns else "—"

        _raw = str(df["PumpPhaseType"].iloc[0]) if "PumpPhaseType" in df.columns else "—"
        stats["pump_phase"] = (
            "Three Phase Pump" if _raw == "1"
            else "Single Phase Pump" if _raw == "0"
            else _raw
        )

        stats["date_range"] = "—"
        stats["first_timestamp"] = "—"
        stats["last_timestamp"] = "—"

        if "QueuedTime-IST" in df.columns:
            ts = parse_ist_column(df["QueuedTime-IST"]).dropna()
            if not ts.empty:
                stats["date_range"] = f"{ts.min().strftime('%d %b %Y')} → {ts.max().strftime('%d %b %Y')}"
                # ═══════════════════════════════════════════════════════
                # BUG FIX 5: Store timestamps in YYYY-MM-DD format for
                # HTML date inputs, alongside display format for the UI.
                # ═══════════════════════════════════════════════════════
                stats["first_timestamp"] = ts.min().strftime('%d %b %Y %H:%M:%S')
                stats["last_timestamp"] = ts.max().strftime('%d %b %Y %H:%M:%S')
                stats["first_timestamp_iso"] = ts.min().strftime('%Y-%m-%dT%H:%M')
                stats["last_timestamp_iso"] = ts.max().strftime('%Y-%m-%dT%H:%M')

        stats["start_count"] = 0
        stats["stop_count"] = 0
        stats["motor_events"] = []
        if "MotorRunningStatus" in df.columns and "QueuedTime-IST" in df.columns:
            raw = df["MotorRunningStatus"].fillna(0).astype(str).str.strip().str.upper()
            ms = raw.isin({"1", "TRUE", "YES", "ON"})
            stats["start_count"] = int((~ms.shift(1, fill_value=True) & ms).sum())
            stats["stop_count"]  = int((ms.shift(1, fill_value=False) & ~ms).sum())

            # Compute motor run events (start → stop pairs)
            try:
                _df = df.copy()
                _df["_ts"] = parse_ist_column(_df["QueuedTime-IST"])
                _df["_ms"] = ms.values
                _df["_err"] = pd.to_numeric(_df.get("Error CondMon", pd.Series(0, index=_df.index)), errors="coerce").fillna(0).astype(int)
                _df = _df.sort_values("_ts").reset_index(drop=True)

                events = []
                in_run = False
                start_ts = None
                run_errors = set()

                for idx, row in _df.iterrows():
                    is_on = bool(row["_ms"])
                    ts_val = row["_ts"]
                    err_val = int(row["_err"])

                    if not in_run and is_on:
                        in_run = True
                        start_ts = ts_val
                        run_errors = set()
                    elif in_run:
                        if err_val != 0:
                            run_errors.add(err_val)
                        if not is_on:
                            stop_ts = ts_val
                            dur_secs = (stop_ts - start_ts).total_seconds() if pd.notna(start_ts) and pd.notna(stop_ts) else 0
                            days = int(dur_secs // 86400)
                            hours = int((dur_secs % 86400) // 3600)
                            mins = int((dur_secs % 3600) // 60)
                            secs = int(dur_secs % 60)
                            dur_str = f"{days}D {hours:02d}H {mins:02d}M {secs:02d}S" if days > 0 else f"{hours:02d}H {mins:02d}M {secs:02d}S"
                            err_descs = [ERROR_CODES.get(e, f"Code {e}") for e in sorted(run_errors)] if run_errors else []
                            events.append({
                                "num": len(events) + 1,
                                "start": start_ts.strftime("%d %b %Y %H:%M:%S") if pd.notna(start_ts) else "—",
                                "stop": stop_ts.strftime("%d %b %Y %H:%M:%S") if pd.notna(stop_ts) else "—",
                                "duration": dur_str,
                                "errors": err_descs,
                                "has_error": len(err_descs) > 0,
                            })
                            in_run = False
                            start_ts = None
                            run_errors = set()

                # If still running at end of data
                if in_run and start_ts is not None:
                    last_ts = _df["_ts"].iloc[-1]
                    dur_secs = (last_ts - start_ts).total_seconds() if pd.notna(last_ts) else 0
                    days = int(dur_secs // 86400)
                    hours = int((dur_secs % 86400) // 3600)
                    mins = int((dur_secs % 3600) // 60)
                    secs = int(dur_secs % 60)
                    dur_str = f"{days}D {hours:02d}H {mins:02d}M {secs:02d}S" if days > 0 else f"{hours:02d}H {mins:02d}M {secs:02d}S"
                    events.append({
                        "num": len(events) + 1,
                        "start": start_ts.strftime("%d %b %Y %H:%M:%S") if pd.notna(start_ts) else "—",
                        "stop": "Still Running ▶",
                        "duration": dur_str,
                        "errors": [],
                        "has_error": False,
                    })

                stats["motor_events"] = events
                logger.info(f"Motor events computed: {len(events)} runs")
            except Exception as e:
                logger.warning(f"Motor events computation failed: {e}")
                stats["motor_events"] = []
        elif "MotorRunningStatus" in df.columns:
            raw = df["MotorRunningStatus"].fillna(0).astype(str).str.strip().str.upper()
            ms = raw.isin({"1", "TRUE", "YES", "ON"})
            stats["start_count"] = int((~ms.shift(1, fill_value=True) & ms).sum())
            stats["stop_count"]  = int((ms.shift(1, fill_value=False) & ~ms).sum())

        stats["error_summary"] = {}
        stats["error_count_by_code"] = {}

        if "Error CondMon" in df.columns:
            codes = pd.to_numeric(df["Error CondMon"], errors="coerce").dropna().astype(int)
            for code, count in codes.value_counts().items():
                try:
                    code_int = int(code)
                    desc = ERROR_CODES.get(code_int, "Unknown Error")
                    severity = (
                        "critical" if code_int in CRITICAL_ERRORS
                        else "warning" if code_int in WARNING_ERRORS
                        else "info"
                    )
                    stats["error_summary"][f"{code_int} - {desc}"] = {
                        "count": int(count),
                        "severity": severity
                    }
                    stats["error_count_by_code"][code_int] = int(count)
                except (ValueError, TypeError) as e:
                    logger.warning(f"Error code processing failed: {e}")
                    continue

        dur = compute_durations(df)

        stats["total_log_mins"] = dur["total_log_mins"]
        stats["motor_idle_mins"] = dur["motor_idle_mins"]
        stats["no_error_running_mins"] = dur["no_error_running_mins"]
        stats["offline_mins"] = dur["offline_mins"]
        stats["error_duration_by_code"] = dur["error_duration_by_code"]
        stats["motor_running_mins"] = dur["motor_running_mins"]

        # ═══════════════════════════════════════════════════════
        # BUG FIX 6: error_running_mins must come from dur (motor ON
        # + error active time only), NOT from sum of error_duration_by_code
        # which also includes motor-OFF error time and would over-count.
        # total_error_mins is the separate metric for all error time
        # (motor ON or OFF) used to display "Error Time" on the dashboard.
        # ═══════════════════════════════════════════════════════
        stats["error_running_mins"] = dur["error_running_mins"]
        stats["total_error_mins"] = sum(dur["error_duration_by_code"].values())

        stats["mode_summary"] = {}
        if "ModeOfOperating" in df.columns:
            for code, count in df["ModeOfOperating"].value_counts().items():
                try:
                    code_int = int(code)
                    mode_desc = MODE_DESCRIPTIONS.get(code_int, f"Unknown Mode ({code})")
                    stats["mode_summary"][mode_desc] = {
                        "code": code_int,
                        "count": int(count),
                        "icon": MODE_ICONS.get(code_int, "❓"),
                        "duration": minutes_to_dhm(dur["mode_runtime"].get(code_int, 0)),
                    }
                except (ValueError, TypeError) as e:
                    logger.warning(f"Mode processing failed: {e}")
                    continue

        stats["error_duration_table"] = []
        for code, mins in sorted(dur["error_duration_by_code"].items()):
            try:
                stats["error_duration_table"].append({
                    "code": code,
                    "desc": ERROR_CODES.get(code, "Unknown Error"),
                    "severity": (
                        "critical" if code in CRITICAL_ERRORS
                        else "warning" if code in WARNING_ERRORS
                        else "info"
                    ),
                    "count": stats["error_count_by_code"].get(code, 0),
                    "duration": minutes_to_dhm(mins),
                    "minutes": round(mins, 1),
                })
            except Exception as e:
                logger.warning(f"Error duration table entry failed: {e}")
                continue

        def _s(col):
            return df[col] if col in df.columns else pd.Series(dtype=float)

        v1_min, v1_max, v1_avg = col_stats(_s("Line Voltage"))
        v2_min, v2_max, v2_avg = col_stats(_s("Line Voltage 2"))
        v3_min, v3_max, v3_avg = col_stats(_s("Line Voltage 3"))

        stats["v1_min"], stats["v1_max"], stats["v1_avg"] = v1_min, v1_max, v1_avg
        stats["v2_min"], stats["v2_max"], stats["v2_avg"] = v2_min, v2_max, v2_avg
        stats["v3_min"], stats["v3_max"], stats["v3_avg"] = v3_min, v3_max, v3_avg
        stats["avg_voltage"] = group_avg(df, VOLTAGE_COLS)
        stats["avg_voltage_no_fail"] = calc_avg_voltage_no_fail(df, VOLTAGE_COLS)

        c1_min, c1_max, c1_avg = col_stats_motor_on(df, "Current Amp")
        c2_min, c2_max, c2_avg = col_stats_motor_on(df, "Current Amp2")
        c3_min, c3_max, c3_avg = col_stats_motor_on(df, "Current Amp3")

        stats["c1_min"], stats["c1_max"], stats["c1_avg"] = c1_min, c1_max, c1_avg
        stats["c2_min"], stats["c2_max"], stats["c2_avg"] = c2_min, c2_max, c2_avg
        stats["c3_min"], stats["c3_max"], stats["c3_avg"] = c3_min, c3_max, c3_avg

        valid_currents = [c for c in [c1_avg, c2_avg, c3_avg] if c > 0]
        stats["avg_current"] = (
            round(sum(valid_currents) / len(valid_currents), 2)
            if valid_currents
            else round((c1_avg + c2_avg + c3_avg) / 3, 2)
        )

        p_min, p_max, p_avg = col_stats(_s("Pressure"))
        f_min, f_max, f_avg = col_stats(_s("Flow Sensor"))
        freq_min, freq_max, freq_avg = col_stats(_s("Frequency"))
        sig_min, sig_max, sig_avg = col_stats(_s("Signal"))

        stats["p_min"], stats["p_max"], stats["p_avg"] = p_min, p_max, p_avg
        stats["f_min"], stats["f_max"], stats["f_avg"] = f_min, f_max, f_avg
        stats["freq_min"], stats["freq_max"], stats["freq_avg"] = freq_min, freq_max, freq_avg
        stats["sig_min"], stats["sig_max"], stats["sig_avg"] = sig_min, sig_max, sig_avg

        # PackCount span + 1 (inclusive: count both first and last packet)
        stats["pack_count_total"] = 0
        if "PackCount" in df.columns:
            try:
                valid_pc = pd.to_numeric(df["PackCount"], errors="coerce").dropna()
                if len(valid_pc) > 0:
                    diff2 = valid_pc.diff().fillna(0)
                    if (diff2 < 0).any():
                        stats["pack_count_total"] = int(diff2[diff2 > 0].sum()) + 1
                    else:
                        stats["pack_count_total"] = int(valid_pc.iloc[-1] - valid_pc.iloc[0]) + 1
            except Exception as e:
                logger.warning(f"Pack count calculation failed: {e}")

        net_4g_count = net_2g_count = 0
        if "NetType" in df.columns:
            is_4g = df["NetType"].astype(str).str.upper() == "4G"
            net_4g_count = int(is_4g.sum())
            net_2g_count = len(df) - net_4g_count

        stats["network_pct_4g"] = round(
            (net_4g_count / max(net_4g_count + net_2g_count, 1)) * 100, 1
        )

        score = 100.0
        for key, val in stats["error_summary"].items():
            try:
                code = int(key.split(" - ")[0])
                if code == 0:
                    continue
                cnt = val["count"]
                if code in CRITICAL_ERRORS:
                    score -= min(15, cnt * 3)
                elif code in WARNING_ERRORS:
                    score -= min(8, cnt)
                else:
                    score -= min(4, cnt * 0.5)
            except (ValueError, IndexError) as e:
                logger.warning(f"Health score calculation error: {e}")
                continue

        stats["health_score"] = max(0, min(100, round(score, 1)))
        stats["error_events"] = compute_error_events(df)

        stats["quality_findings"] = []
        checks = [
            ("Line Voltage", "R-Phase Voltage"),
            ("Line Voltage 2", "Y-Phase Voltage"),
            ("Line Voltage 3", "B-Phase Voltage"),
            ("Current Amp", "R-Phase Current"),
            ("Current Amp2", "Y-Phase Current"),
            ("Current Amp3", "B-Phase Current"),
            ("Pressure", "Pressure"),
            ("Flow Sensor", "Flow Sensor"),
            ("Frequency", "Frequency"),
        ]

        for col, label in checks:
            if col in df.columns:
                data = pd.to_numeric(df[col], errors="coerce").dropna()
                if len(data) == 0:
                    stats["quality_findings"].append({"label": label, "status": "missing", "message": "No data recorded"})
                elif (data == 0).all():
                    stats["quality_findings"].append({"label": label, "status": "warn", "message": "All values are zero"})
                else:
                    stats["quality_findings"].append({"label": label, "status": "ok", "message": f"{len(data)} readings"})
            else:
                stats["quality_findings"].append({"label": label, "status": "missing", "message": "Column not found"})

        logger.info("Statistics computation complete")
        return stats

    except Exception as e:
        logger.error(f"Statistics computation failed: {e}", exc_info=True)
        return {
            "error": str(e),
            "health_score": 0,
            "error_summary": {},
            "quality_findings": []
        }


def compute_charts(df):
    charts = {}
    try:
        if "QueuedTime-IST" in df.columns:
            df = df.copy()
            df["QueuedTime-IST"] = parse_ist_column(df["QueuedTime-IST"])
            labels = [t.strftime("%d-%m-%Y %H:%M:%S") if pd.notna(t) else "—"
                      for t in df["QueuedTime-IST"]]
        else:
            labels = [str(i) for i in range(len(df))]

        charts["labels"] = downsample_list(labels)

        chart_data = [
            ("voltage1", "Line Voltage"),
            ("voltage2", "Line Voltage 2"),
            ("voltage3", "Line Voltage 3"),
            ("current1", "Current Amp"),
            ("current2", "Current Amp2"),
            ("current3", "Current Amp3"),
            ("pressure", "Pressure"),
            ("flow", "Flow Sensor"),
            ("frequency", "Frequency"),
            ("signal", "Signal"),
            ("pack_count_graph", "PackCount"),
        ]

        for key, col in chart_data:
            charts[key] = downsample_series(df[col]) if col in df.columns else []

        charts["network_numeric"] = []
        if "NetType" in df.columns:
            nettype = df["NetType"].astype(str).str.upper().map({"4G": 4, "2G": 2})
            charts["network_numeric"] = downsample_series(nettype.fillna(1))

        # BUG FIX 7 (continued): cast to plain Python bool — numpy.bool_ is not JSON serializable
        charts["has_voltage"]   = bool(all(col in df.columns for col in VOLTAGE_COLS) and any(df[VOLTAGE_COLS].notna().any()))
        charts["has_current"]   = bool(all(col in df.columns for col in CURRENT_COLS) and any(df[CURRENT_COLS].notna().any()))
        charts["has_pressure"]  = bool("Pressure" in df.columns and df["Pressure"].notna().any())
        charts["has_flow"]      = bool("Flow Sensor" in df.columns and df["Flow Sensor"].notna().any())
        charts["has_frequency"] = bool("Frequency" in df.columns and df["Frequency"].notna().any())
        charts["has_signal"]    = bool("Signal" in df.columns and df["Signal"].notna().any())
        charts["has_pack"]      = bool("PackCount" in df.columns and df["PackCount"].notna().any())
        charts["has_network"]   = bool("NetType" in df.columns and df["NetType"].notna().any())

        logger.info("Chart data prepared")
        return charts

    except Exception as e:
        logger.error(f"Chart computation failed: {e}", exc_info=True)
        return {
            "labels": [],
            "has_voltage": False, "has_current": False, "has_pressure": False,
            "has_flow": False, "has_frequency": False, "has_signal": False,
            "has_pack": False, "has_network": False,
        }


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/progress", methods=["GET"])
def get_progress():
    return jsonify(progress_data)


@app.route("/upload", methods=["POST"])
def upload():
    set_progress(0, "Initializing...")
    try:
        file = request.files.get("file")
        is_valid, message = validate_upload(file)
        if not is_valid:
            logger.warning(f"Upload validation failed: {message}")
            set_progress(0, "Idle")
            return render_template("index.html", upload_error=message)

        set_progress(10, "Reading file...")

        try:
            file_bytes = file.read()
            file.seek(0)
        except Exception as e:
            set_progress(0, "Idle")
            logger.error(f"File read error: {e}")
            return render_template("index.html", upload_error=f"Could not read file: {e}")

        filename = file.filename
        engine = "openpyxl" if filename.endswith(".xlsx") else "xlrd"

        tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        tmp_file.write(file_bytes)
        tmp_file.close()

        session["tmp_path"] = tmp_file.name
        session["filename"] = filename
        session["engine"] = engine

        set_progress(30, "Reading data...")
        try:
            buf = io.BytesIO(file_bytes)
            df1 = read_excel_cols(buf, engine, STATS_COLS)
            set_progress(50, "Calculating statistics...")
            stats = compute_stats(df1)
            del df1
            gc.collect()
        except Exception as e:
            set_progress(0, "Idle")
            logger.error(f"Stats computation error: {e}")
            return render_template("index.html", upload_error=f"Stats error: {e}")

        set_progress(70, "Generating charts...")
        try:
            buf.seek(0)
            df2 = read_excel_cols(buf, engine, CHART_COLS)
            charts = compute_charts(df2)
            del df2, buf
            gc.collect()
        except Exception as e:
            set_progress(0, "Idle")
            logger.error(f"Chart computation error: {e}")
            return render_template("index.html", upload_error=f"Chart error: {e}")

        set_progress(90, "Building dashboard...")

        stats["pack_count_total"] = int(stats.get("pack_count_total", 0))
        stats["running_time"] = minutes_to_dhm(stats.get("motor_running_mins", 0))
        stats["total_log_time"] = minutes_to_dhm(stats.get("total_log_mins", 0))
        stats["offline_time"] = minutes_to_dhm(stats.get("offline_mins", 0))
        stats["idle_time"] = minutes_to_dhm(stats.get("motor_idle_mins", 0))
        # ═══════════════════════════════════════════════════════
        # BUG FIX 6 (continued): Use total_error_mins (motor ON + OFF)
        # for the dashboard "Error Time" display card, not error_running_mins.
        # ═══════════════════════════════════════════════════════
        stats["error_time"] = minutes_to_dhm(stats.get("total_error_mins", 0))
        stats["no_error_run_time"] = minutes_to_dhm(stats.get("no_error_running_mins", 0))

        hs = stats.get("health_score", 0)
        stats["health_label"] = (
            "Excellent" if hs == 100
            else "Good" if hs >= 85
            else "Fair" if hs >= 50
            else "Poor"
        )
        stats["health_color"] = (
            "green" if hs == 100
            else "blue" if hs >= 85
            else "gold" if hs >= 50
            else "red"
        )

        error_summary = stats.get("error_summary", {})
        stats["critical_count"] = sum(v["count"] for v in error_summary.values()
                                      if v.get("severity") == "critical")
        stats["warning_count"] = sum(v["count"] for v in error_summary.values()
                                     if v.get("severity") == "warning")
        stats["total_errors"] = sum(v["count"] for k, v in error_summary.items()
                                    if not k.startswith("0 -"))

        summary_points = []
        if stats["critical_count"] > 0:
            summary_points.append({"type": "critical", "icon": "🔴",
                                    "text": f"{stats['critical_count']} critical fault event(s) detected."})
        if stats["warning_count"] > 0:
            summary_points.append({"type": "warning", "icon": "🟡",
                                    "text": f"{stats['warning_count']} warning event(s) logged."})
        if stats["total_errors"] == 0:
            summary_points.append({"type": "ok", "icon": "✅", "text": "Zero fault conditions recorded."})
        if hs >= 85:
            summary_points.append({"type": "ok", "icon": "🟢",
                                    "text": f"Excellent condition. Health score: {hs}."})

        stats["summary_points"] = summary_points

        quality_findings = stats.get("quality_findings", [])
        stats["quality_ok"] = sum(1 for f in quality_findings if f["status"] == "ok")
        stats["quality_warn"] = sum(1 for f in quality_findings if f["status"] == "warn")
        stats["quality_miss"] = sum(1 for f in quality_findings if f["status"] == "missing")

        set_progress(100, "Dashboard Ready!")

        # ═══════════════════════════════════════════════════════
        # BUG FIX 5 (continued): Use ISO format timestamps for
        # session and date filter inputs so HTML date pickers work.
        # ═══════════════════════════════════════════════════════
        min_date = stats.get("first_timestamp_iso", stats.get("first_timestamp", ""))
        max_date = stats.get("last_timestamp_iso", stats.get("last_timestamp", ""))
        session["min_date"] = min_date
        session["max_date"] = max_date

        return render_template(
            "dashboard.html",
            filename=filename,
            filter_applied=False,
            filter_from=min_date,
            filter_to=max_date,
            data_min_date=min_date,
            data_max_date=max_date,
            **stats,
            **charts,
        )

    except Exception as e:
        set_progress(0, "Idle")
        logger.error(f"Upload processing failed: {e}", exc_info=True)
        return render_template("index.html", upload_error=f"Processing error: {str(e)}")


@app.route("/filter", methods=["POST"])
def filter_data():
    """
    Re-run stats and charts on the uploaded file, restricted to the
    date range submitted by the dashboard filter sidebar.
    """
    set_progress(0, "Applying filter...")
    try:
        tmp_path = session.get("tmp_path")
        filename = session.get("filename", "report.xlsx")
        engine   = session.get("engine", "openpyxl")

        if not tmp_path or not os.path.exists(tmp_path):
            logger.warning("Filter requested but no uploaded file in session")
            return render_template("index.html",
                                   upload_error="Session expired. Please re-upload your file.")

        filter_type = request.form.get("filter_type", "range")
        date_from   = request.form.get("date_from", "").strip()
        date_to     = request.form.get("date_to", "").strip()
        filter_applied = filter_type == "range" and bool(date_from and date_to)

        set_progress(20, "Reading data...")
        try:
            with open(tmp_path, "rb") as f:
                file_bytes = f.read()
        except Exception as e:
            logger.error(f"Could not read temp file: {e}")
            return render_template("index.html", upload_error=f"Could not read uploaded file: {e}")

        buf = io.BytesIO(file_bytes)

        set_progress(35, "Loading columns...")
        df_full = read_excel_cols(buf, engine, STATS_COLS + [c for c in CHART_COLS if c not in STATS_COLS])

        # Apply date filter
        if filter_applied and "QueuedTime-IST" in df_full.columns:
            try:
                df_full["QueuedTime-IST"] = parse_ist_column(df_full["QueuedTime-IST"])
                ts_from = pd.to_datetime(date_from, errors="coerce")
                ts_to   = pd.to_datetime(date_to,   errors="coerce")
                if pd.notna(ts_from) and pd.notna(ts_to):
                    # Make ts_to inclusive for the full day
                    ts_to = ts_to + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
                    mask = (df_full["QueuedTime-IST"] >= ts_from) & (df_full["QueuedTime-IST"] <= ts_to)
                    df_filtered = df_full[mask].reset_index(drop=True)
                    logger.info(f"Filter applied: {ts_from} → {ts_to}, {len(df_filtered)} rows")
                    if df_filtered.empty:
                        logger.warning("Filter returned 0 rows — showing full dataset")
                        df_filtered = df_full
                        filter_applied = False
                else:
                    df_filtered = df_full
                    filter_applied = False
            except Exception as e:
                logger.warning(f"Date filter failed: {e}")
                df_filtered = df_full
                filter_applied = False
        else:
            df_filtered = df_full
            filter_applied = False

        set_progress(50, "Calculating statistics...")
        stats = compute_stats(df_filtered[STATS_COLS if all(c in df_filtered.columns for c in STATS_COLS)
                                           else [c for c in STATS_COLS if c in df_filtered.columns]])

        set_progress(70, "Generating charts...")
        charts = compute_charts(df_filtered[[c for c in CHART_COLS if c in df_filtered.columns]])

        del df_full, df_filtered, buf
        gc.collect()

        set_progress(90, "Building dashboard...")

        stats["pack_count_total"] = int(stats.get("pack_count_total", 0))
        stats["running_time"]     = minutes_to_dhm(stats.get("motor_running_mins", 0))
        stats["total_log_time"]   = minutes_to_dhm(stats.get("total_log_mins", 0))
        stats["offline_time"]     = minutes_to_dhm(stats.get("offline_mins", 0))
        stats["idle_time"]        = minutes_to_dhm(stats.get("motor_idle_mins", 0))
        stats["error_time"]       = minutes_to_dhm(stats.get("total_error_mins", 0))
        stats["no_error_run_time"]= minutes_to_dhm(stats.get("no_error_running_mins", 0))

        hs = stats.get("health_score", 0)
        stats["health_label"] = (
            "Excellent" if hs == 100 else "Good" if hs >= 85 else "Fair" if hs >= 50 else "Poor"
        )
        stats["health_color"] = (
            "green" if hs == 100 else "blue" if hs >= 85 else "gold" if hs >= 50 else "red"
        )

        error_summary = stats.get("error_summary", {})
        stats["critical_count"] = sum(v["count"] for v in error_summary.values() if v.get("severity") == "critical")
        stats["warning_count"]  = sum(v["count"] for v in error_summary.values() if v.get("severity") == "warning")
        stats["total_errors"]   = sum(v["count"] for k, v in error_summary.items() if not k.startswith("0 -"))

        summary_points = []
        if stats["critical_count"] > 0:
            summary_points.append({"type": "critical", "icon": "🔴",
                                    "text": f"{stats['critical_count']} critical fault event(s) detected."})
        if stats["warning_count"] > 0:
            summary_points.append({"type": "warning", "icon": "🟡",
                                    "text": f"{stats['warning_count']} warning event(s) logged."})
        if stats["total_errors"] == 0:
            summary_points.append({"type": "ok", "icon": "✅", "text": "Zero fault conditions recorded."})
        if hs >= 85:
            summary_points.append({"type": "ok", "icon": "🟢",
                                    "text": f"Excellent condition. Health score: {hs}."})
        stats["summary_points"] = summary_points

        quality_findings = stats.get("quality_findings", [])
        stats["quality_ok"]   = sum(1 for f in quality_findings if f["status"] == "ok")
        stats["quality_warn"] = sum(1 for f in quality_findings if f["status"] == "warn")
        stats["quality_miss"] = sum(1 for f in quality_findings if f["status"] == "missing")

        data_min = session.get("min_date", "")
        data_max = session.get("max_date", "")

        set_progress(100, "Dashboard Ready!")

        return render_template(
            "dashboard.html",
            filename=filename,
            filter_applied=filter_applied,
            filter_from=date_from if filter_applied else data_min,
            filter_to=date_to   if filter_applied else data_max,
            data_min_date=data_min,
            data_max_date=data_max,
            **stats,
            **charts,
        )

    except Exception as e:
        set_progress(0, "Idle")
        logger.error(f"Filter processing failed: {e}", exc_info=True)
        return render_template("index.html", upload_error=f"Filter error: {str(e)}")


if __name__ == "__main__":
    logger.info("Starting TARO CONNECT Application")
    app.run(debug=False)
