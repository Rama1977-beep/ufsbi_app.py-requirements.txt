import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

st.set_page_config(page_title="UFSBI Expert Analyzer V3", layout="wide")

# ==========================================================
# UFSBI EXPERT ANALYZER V3
# New in V3:
# 1. It does NOT search the whole file blindly.
# 2. It detects the operation start point automatically.
# 3. It checks relay sequence only after the detected start point.
# 4. Link Error is declared only if latest/final BIPR1 and BIPR2 status are DOWN.
# 5. B1PR1 / B1PR2 / BPR1 / BPR2 / FR1 / FR2 are NOT used for automatic Telecom failure.
# ==========================================================

MASTER_SEQUENCE = [
    [1, "Line Clear Initiating", "BPNR", "UP", "BPNR not pickup: check bell button contact and SM key."],
    [2, "Line Clear Initiating", "TGTNR", "UP", "TGTNR not pickup: check TGT button and CNR relay contact D7/D8."],
    [3, "Line Clear Initiating", "TGTXR", "UP", "TGTXR not pickup: check TGTR drop A7/A8, BPNR B1/N2, TGTNR A1/A2, ASGNCR A1/A2, DAZTR B1/B2, 24V fuse."],
    [4, "Line Clear Link", "TGTYR", "UP", "TGTYR not pickup: check TGTYR R1/R2. If BIPR1/BIPR2 latest status down, inform Telecom."],
    [5, "Train On Line Prep", "HSGNCR", "DOWN", "HSGNCR not drop: check HSGNCR drop path."],
    [6, "Train On Line Prep", "HSATPR", "DOWN", "HSATPR not drop: check HSATPR drop proving."],
    [7, "Train On Line Prep", "TAR1", "UP", "TAR1 not pickup/hold: check HSATPR A7/A8, TAR1 A2/A1, HSBTPR D2/D1, TCFR A3/A4, TAR2 D5/D6."],
    [8, "Train On Line Prep", "HSBTPR", "DOWN", "HSBTPR not drop: check HSBTPR drop path."],
    [9, "Train On Line Prep", "HSATPR", "UP", "HSATPR not pickup: check BTSR A5/A6, HSBTPR D7/D8, HSATPR A1/A2, TAR1 B1/N2, TAR2 D2/D1."],
    [10, "Train On Line Prep", "TAR2", "UP", "TAR2 not pickup/hold: check BTSR A5/A6, HSBTPR D7/D8, HSATPR A1/A2, TAR1 B1/N2, TAR2 D2/D1."],
    [11, "Train On Line Prep", "TAR1", "DOWN", "TAR1 not drop: check TAR1 drop circuit/holding path."],
    [12, "Train On Line", "DAZTR", "UP", "DAZTR not pickup: check DAZTR pickup circuit and 24V feed."],
    [13, "Train On Line", "HSGNCR", "UP", "HSGNCR not pickup: check HSGNCR pickup circuit."],
    [14, "Train On Line", "TCFR", "DOWN", "TCFR not drop: check TCFXR B7/B8, TAR2 C2/C1, CAR D8/D7, ASGNCPR A1/A2, CNR D6/D5, HSGNCR D1/D2, DAZTR C1/C2, LCB key reverse."],
    [15, "Train On Line", "BTSR", "UP", "BTSR not pickup: check RAZTR C4/C5, CAR drop D5/D6, TCFR drop D7/D8, BTSR pickup A2/A1."],
    [16, "Train On Line Link", "TGTNZR", "DOWN", "TGTNZR not drop: check STN-B."],
    [17, "Train On Line", "TGTR", "UP", "TGTR not pickup: check TGTR R1/R2 voltage and previous relay contacts."],
    [18, "Arrival / Proving", "DAZTR", "UP", "Check DAZTR contact B1/N2."],
    [19, "Arrival / Proving", "ASGNCR", "UP", "Check ASGNCR contact A1/A2."],
    [20, "Arrival / Proving", "BPNR", "UP", "Check BPNR contact B1/N2."],
    [21, "Arrival / Proving", "TGTYR", "UP", "Check TGTYR contact A1/A2."],
    [22, "Arrival / Proving", "TGTXR", "UP", "Check TGTXR contact A1/D2."]
]

START_RELAYS = [
    ("BPNR", "UP", 1),
    ("TGTNR", "UP", 2),
    ("TGTXR", "UP", 3),
    ("TGTYR", "UP", 4),
]

def safe(x):
    try:
        if pd.isna(x):
            return ""
    except Exception:
        pass
    return str(x)

def clean_relay(x):
    t = safe(x).upper().strip()
    t = re.sub(r"\(.*?\)", "", t)
    t = re.sub(r"\s+", "", t)
    return t

def clean_status(x):
    t = safe(x).upper().strip()
    if t in ["PICKUP", "PICK UP", "PICK", "PU", "ON"]:
        return "UP"
    if t in ["DROP", "DROPPED", "OFF"]:
        return "DOWN"
    return t

def read_excel(file):
    errors = []
    for eng in ["openpyxl", "xlrd", None]:
        try:
            file.seek(0)
            if eng:
                return pd.read_excel(file, engine=eng, header=None, dtype=str)
            return pd.read_excel(file, header=None, dtype=str)
        except Exception as e:
            errors.append(f"{eng}: {e}")
    raise Exception("Excel file read नहीं हो रहा. Details: " + " | ".join(errors))

def relay_score_value(v):
    v = safe(v).upper().strip()
    if not v or v in ["UP", "DOWN", "ON", "OFF", "NAN"]:
        return 0
    if "(" in v and ")" in v:
        return 5
    if re.search(r"[A-Z]{2,}", v) and not re.search(r"\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4}", v):
        return 1
    return 0

def time_score_value(v):
    v = safe(v)
    if re.search(r"\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4}.*\d{1,2}:\d{2}", v):
        return 3
    if re.search(r"\d{1,2}:\d{2}:\d{2}", v):
        return 1
    return 0

def detect_columns(df):
    relay_col = status_col = time_col = None
    best_relay = best_status = best_time = -1

    for c in df.columns:
        vals = df[c].fillna("").astype(str).head(1500)
        relay_score = sum(relay_score_value(v) for v in vals)
        status_score = sum(clean_status(v) in ["UP", "DOWN"] for v in vals)
        time_score = sum(time_score_value(v) for v in vals)

        if relay_score > best_relay:
            best_relay, relay_col = relay_score, c
        if status_score > best_status:
            best_status, status_col = status_score, c
        if time_score > best_time:
            best_time, time_col = time_score, c

    return relay_col, status_col, time_col

def extract_events(df):
    rc, sc, tc = detect_columns(df)
    if rc is None or sc is None:
        raise Exception("Relay/Status column detect नहीं हुआ.")

    events = []
    for i, row in df.iterrows():
        relay = clean_relay(row[rc])
        status = clean_status(row[sc])
        if relay and status in ["UP", "DOWN"]:
            events.append({
                "index": len(events),
                "row": i + 1,
                "relay": relay,
                "status": status,
                "time": safe(row[tc]) if tc is not None else "-"
            })
    return events, {"relay": rc, "status": sc, "time": tc}

def find_next_event(events, relay, status, start_index):
    for i in range(start_index, len(events)):
        if events[i]["relay"] == relay and events[i]["status"] == status:
            return i, events[i]
    return None, None

def latest_status(events, relay):
    for e in reversed(events):
        if e["relay"] == relay:
            return e["status"], e
    return None, None

def detect_link_failure(events):
    """
    Link Error is declared only if latest/final actual BIPR1 and BIPR2
    status are both DOWN.
    This avoids false link errors when BIPR1/BIPR2 drop/pick during normal log.
    """
    bipr1_status, bipr1_event = latest_status(events, "BIPR1")
    bipr2_status, bipr2_event = latest_status(events, "BIPR2")

    if bipr1_status == "DOWN" and bipr2_status == "DOWN":
        return {
            "failure_type": "LINK ERROR / COMMUNICATION FAILURE",
            "department": "TELECOM",
            "reason": f"Latest status: BIPR1 DOWN at row {bipr1_event['row']} and BIPR2 DOWN at row {bipr2_event['row']}.",
            "action": "Inform Telecom. Check OFC link, modem, media converter, quad cable and communication network. Relay contact checking not required."
        }

    return None

def classify(step):
    if step <= 4:
        return "Line Clear Not Initiating / Link Response Failure"
    if step <= 11:
        return "Train On Line Initiation Not Completed"
    if step <= 17:
        return "Train On Line Not Completed"
    return "Arrival / Proving Not Completed"

def evaluate_candidate(events, candidate_event_index, start_step):
    """
    Start sequence from detected operation start.
    If log starts from middle of operation, start_step can be 2/3/4.
    """
    seq_index = start_step - 1
    pos = candidate_event_index
    matched = []
    rows = []
    fail = None

    for rule in MASTER_SEQUENCE[seq_index:]:
        step, phase, relay, status, check = rule
        found_index, ev = find_next_event(events, relay, status, pos)

        if ev:
            rows.append([step, phase, relay, status, ev["time"], ev["row"], "OK", "-"])
            matched.append((step, found_index))
            pos = found_index + 1
        else:
            rows.append([step, phase, relay, status, "-", "-", "MISSING", check])
            fail = {
                "step": step,
                "phase": phase,
                "relay": relay,
                "status": status,
                "check": check,
                "failure_type": classify(step),
                "department": "SIGNAL"
            }
            break

    return {
        "start_step": start_step,
        "start_event_index": candidate_event_index,
        "matched_count": len(matched),
        "rows": rows,
        "fail": fail
    }

def detect_operation_and_analyze(events):
    candidates = []

    for relay, status, start_step in START_RELAYS:
        for i, ev in enumerate(events):
            if ev["relay"] == relay and ev["status"] == status:
                result = evaluate_candidate(events, i, start_step)
                candidates.append(result)

    if not candidates:
        fail = {
            "step": 0,
            "phase": "Operation Start",
            "relay": "BPNR/TGTNR/TGTXR/TGTYR",
            "status": "UP",
            "check": "Operation start not found. Check whether selected file contains the actual failure time window.",
            "failure_type": "Operation Start Not Found",
            "department": "DATA / LOG SELECTION"
        }
        rows = [[0, "Operation Start", "BPNR/TGTNR/TGTXR/TGTYR", "UP", "-", "-", "NOT FOUND", fail["check"]]]
        return rows, fail, None

    # Choose candidate with maximum matched sequence.
    # If tie, choose the latest candidate because failure log may contain earlier normal cycles.
    candidates.sort(key=lambda x: (x["matched_count"], x["start_event_index"]), reverse=True)
    best = candidates[0]

    # Add previous skipped note when operation started from step 2/3/4
    if best["start_step"] > 1:
        note = [
            f"Started from Step {best['start_step']}",
            "Auto Operation Detection",
            "Previous steps",
            "SKIPPED",
            "-",
            "-",
            "INFO",
            "Log appears to start from middle of operation. Earlier steps may be outside selected Data Logger time window."
        ]
        best["rows"].insert(0, note)

    return best["rows"], best["fail"], best

def html(rows, cols):
    h = "<table style='border-collapse:collapse;width:100%;font-size:13px'><tr>"
    h += "".join(f"<th style='border:1px solid #aaa;padding:6px;background:#eee'>{c}</th>" for c in cols)
    h += "</tr>"
    for r in rows:
        h += "<tr>"
        h += "".join(f"<td style='border:1px solid #aaa;padding:6px;vertical-align:top'>{x}</td>" for x in r)
        h += "</tr>"
    h += "</table>"
    return h

def make_pdf(rows, cols, failure, fname, meta, operation_info):
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), leftMargin=18, rightMargin=18, topMargin=18, bottomMargin=18)
    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph("<b>UFSBI Expert Failure Analysis Report V3</b>", styles["Title"]))
    story.append(Spacer(1, 8))
    story.append(Paragraph(f"<b>Generated:</b> {datetime.now().strftime('%d-%m-%Y %H:%M:%S')}", styles["Normal"]))
    story.append(Paragraph(f"<b>File:</b> {fname}", styles["Normal"]))

    if operation_info:
        story.append(Paragraph(f"<b>Operation Auto Detected:</b> Start step {operation_info['start_step']}, matched events {operation_info['matched_count']}", styles["Normal"]))

    if failure:
        story.append(Paragraph(f"<b>Failure Type:</b> {failure.get('failure_type', '-')}", styles["Normal"]))
        story.append(Paragraph(f"<b>Department:</b> {failure.get('department', '-')}", styles["Normal"]))
        if "reason" in failure:
            story.append(Paragraph(f"<b>Reason:</b> {failure.get('reason', '-')}", styles["Normal"]))

    story.append(Paragraph(f"<b>Detected Columns:</b> relay={meta['relay']}, status={meta['status']}, time={meta['time']}", styles["Normal"]))
    story.append(Spacer(1, 8))

    data = [cols] + [[str(x) for x in r] for r in rows]
    table = Table(data, repeatRows=1, colWidths=[50, 110, 60, 50, 110, 40, 60, 330])
    table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.35, colors.black),
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("FONTSIZE", (0, 0), (-1, -1), 6.2),
        ("VALIGN", (0, 0), (-1, -1), "TOP")
    ]))
    story.append(table)

    if failure:
        story.append(Spacer(1, 10))
        story.append(Paragraph("<b>Failure Pinpointed</b>", styles["Heading2"]))
        if failure.get("department") == "TELECOM":
            story.append(Paragraph(f"<b>Type:</b> {failure['failure_type']}", styles["Normal"]))
            story.append(Paragraph(f"<b>Reason:</b> {failure['reason']}", styles["Normal"]))
            story.append(Paragraph(f"<b>Action:</b> {failure['action']}", styles["Normal"]))
        else:
            story.append(Paragraph(f"<b>Step:</b> {failure['step']}", styles["Normal"]))
            story.append(Paragraph(f"<b>Missing:</b> {failure['relay']} {failure['status']}", styles["Normal"]))
            story.append(Paragraph(f"<b>Maintainer Action:</b> {failure['check']}", styles["Normal"]))

    doc.build(story)
    buf.seek(0)
    return buf

st.title("UFSBI Expert Analyzer V3")
st.write("Faulty UFSBI Data Logger Excel upload करें. Code operation start auto-detect करके sequence break बताएगा.")

file = st.file_uploader("Upload Faulty Excel", type=["xls", "xlsx"])

if file:
    try:
        df = read_excel(file)
        events, meta = extract_events(df)

        cols = ["Step", "Phase", "Relay", "Expected", "Time", "Row", "Result", "Maintainer Check"]

        link_fail = detect_link_failure(events)

        if link_fail:
            rows = [[0, "Communication / Link", "BIPR1/BIPR2", "LATEST DOWN", "-", "-", "LINK ERROR", link_fail["action"]]]
            fail = link_fail
            operation_info = None
            st.error(f"{link_fail['failure_type']} | Department: TELECOM")
            st.write(link_fail["reason"])
            st.write(link_fail["action"])

        else:
            rows, fail, operation_info = detect_operation_and_analyze(events)

            if operation_info:
                st.info(f"Operation start auto-detected from Step {operation_info['start_step']} | Matched events: {operation_info['matched_count']}")

            if fail:
                st.error(f"{fail['failure_type']}: Step {fail['step']} missing - {fail['relay']} {fail['status']}")
                st.write(fail["check"])
            else:
                st.success("Detected operation sequence completed. No sequence mismatch found.")

        st.subheader("Analysis Result")
        st.markdown(html(rows, cols), unsafe_allow_html=True)

        with st.expander("Show first 100 extracted relay events"):
            preview = pd.DataFrame(events[:100])
            st.dataframe(preview.astype(str), use_container_width=True)

        pdf = make_pdf(rows, cols, fail, file.name, meta, operation_info)
        st.download_button("Download PDF Report", pdf, "UFSBI_Expert_Failure_Report_V3.pdf", "application/pdf")

    except Exception as e:
        st.error(str(e))
else:
    st.info("Faulty UFSBI Excel upload करें.")
