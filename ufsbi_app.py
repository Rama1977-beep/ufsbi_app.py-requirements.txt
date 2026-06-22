import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

st.set_page_config(page_title="UFSBI Expert Analyzer V2", layout="wide")

EXPERT_SEQUENCE = [
    [1, "Line Clear Initiating", "BPNR", "UP", "BPNR not pickup: check bell button contact and SM key."],
    [2, "Line Clear Initiating", "TGTNR", "UP", "TGTNR not pickup: check TGT button and CNR relay contact D7/D8."],
    [3, "Line Clear Initiating", "TGTXR", "UP", "TGTXR not pickup: check TGTR drop A7/A8, BPNR B1/N2, TGTNR A1/A2, ASGNCR A1/A2, DAZTR B1/B2, 24V fuse."],
    [4, "Line Clear Link", "TGTYR", "UP", "TGTYR not pickup: link response not received. If BIPR/FR relays indicate link failure, inform Telecom."],
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
    [15, "Train On Line", "BTSR", "UP", "BTSR not pickup: check RAZTR C4/C5, CAR D5/D6, TCFR D7/D8, BTSR A2/A1."],
    [16, "Train On Line Link", "TGTNZR", "DOWN", "TGTNZR not drop: check STN-B."],
    [17, "Train On Line", "TGTR", "UP", "TGTR not pickup: check TGTR R1/R2 voltage and previous relay contacts."],
    [18, "Proving", "DAZTR", "UP", "Check DAZTR contact B1/N2."],
    [19, "Proving", "ASGNCR", "UP", "Check ASGNCR contact A1/A2."],
    [20, "Proving", "BPNR", "UP", "Check BPNR contact B1/N2."],
    [21, "Proving", "TGTYR", "UP", "Check TGTYR contact A1/A2."],
    [22, "Proving", "TGTXR", "UP", "Check TGTXR contact A1/D2."]
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
    return re.sub(r"\s+", "", t)

def clean_status(x):
    t = safe(x).upper().strip()
    if t in ["PICKUP", "PICK UP", "PICK", "PU", "ON"]:
        return "UP"
    if t in ["DROP", "DROPPED", "OFF"]:
        return "DOWN"
    return t

def read_excel(file):
    for eng in ["openpyxl", "xlrd", None]:
        try:
            file.seek(0)
            if eng:
                return pd.read_excel(file, engine=eng, header=None, dtype=str)
            return pd.read_excel(file, header=None, dtype=str)
        except Exception:
            pass
    raise Exception("Excel file read नहीं हो रहा.")

def detect_columns(df):
    relay_col = status_col = time_col = None
    best_relay = best_status = best_time = -1

    for c in df.columns:
        vals = df[c].fillna("").astype(str).head(1000)

        relay_score = sum(
            5 if "(" in v and ")" in v
            else 1 if re.search(r"[A-Z]{2,}", v.upper())
            else 0
            for v in vals
        )

        status_score = sum(clean_status(v) in ["UP", "DOWN"] for v in vals)

        time_score = sum(
            bool(re.search(r"\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4}.*\d{1,2}:\d{2}", v))
            for v in vals
        )

        if relay_score > best_relay:
            best_relay, relay_col = relay_score, c
        if status_score > best_status:
            best_status, status_col = status_score, c
        if time_score > best_time:
            best_time, time_col = time_score, c

    return relay_col, status_col, time_col

def extract_events(df):
    rc, sc, tc = detect_columns(df)
    events = []

    if rc is None or sc is None:
        raise Exception("Relay/Status column detect नहीं हुआ.")

    for i, row in df.iterrows():
        relay = clean_relay(row[rc])
        status = clean_status(row[sc])

        if relay and status in ["UP", "DOWN"]:
            events.append({
                "row": i + 1,
                "relay": relay,
                "status": status,
                "time": safe(row[tc]) if tc is not None else "-"
            })

    return events, {"relay": rc, "status": sc, "time": tc}

def has_event(events, relay, status=None):
    for e in events:
        if e["relay"] == relay:
            if status is None or e["status"] == status:
                return True
    return False

def find_event(events, relay, status, start):
    for i in range(start, len(events)):
        if events[i]["relay"] == relay and events[i]["status"] == status:
            return i, events[i]
    return None, None

def detect_link_failure(events):
    """
    Field logic:
    Link Error / Communication Failure is Telecom side.
    No relay contact checking required for link error.
    """

    bipr1_down = has_event(events, "BIPR1", "DOWN") or has_event(events, "B1PR1", "DOWN")
    bipr2_down = has_event(events, "BIPR2", "DOWN") or has_event(events, "B1PR2", "DOWN")

    fr1_down = has_event(events, "FR1", "DOWN")
    fr2_down = has_event(events, "FR2", "DOWN")

    tgtxr_up = has_event(events, "TGTXR", "UP")
    tgtyr_down = has_event(events, "TGTYR", "DOWN")
    tgtyr_missing = not has_event(events, "TGTYR", "UP")

    if bipr1_down and bipr2_down:
        return {
            "failure_type": "LINK ERROR / COMMUNICATION FAILURE",
            "department": "TELECOM",
            "reason": "BIPR1/BIPR2 found DOWN in Data Logger.",
            "action": "Inform Telecom. Check OFC link, modem, quad cable, media converter and telecom network path. Relay contact checking not required."
        }

    if fr1_down and fr2_down and bipr1_down and bipr2_down:
        return {
            "failure_type": "COMPLETE UFSBI COMMUNICATION FAILURE",
            "department": "TELECOM",
            "reason": "BIPR1/BIPR2 and FR1/FR2 found DOWN.",
            "action": "Inform Telecom. Check OFC/Quad communication, modem, UFSBI communication cards and telecom path. Relay contact checking not required."
        }

    if tgtxr_up and tgtyr_down:
        return {
            "failure_type": "LINK RESPONSE FAILURE",
            "department": "TELECOM",
            "reason": "TGTXR UP but TGTYR DOWN. Local transmission available but remote response not received.",
            "action": "Inform Telecom. Check remote station communication, OFC/Quad link and modem. Relay contact checking not required."
        }

    if tgtxr_up and tgtyr_missing:
        return {
            "failure_type": "POSSIBLE LINK RESPONSE FAILURE",
            "department": "TELECOM",
            "reason": "TGTXR UP found but TGTYR UP not found in expected log.",
            "action": "Verify with opposite station. If TGTYR not received, inform Telecom for OFC/Quad/Modem checking."
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

def analyze_signal_sequence(events):
    rows = []
    pos = 0
    fail = None

    for step, phase, relay, status, check in EXPERT_SEQUENCE:
        idx, ev = find_event(events, relay, status, pos)

        if ev:
            rows.append([step, phase, relay, status, ev["time"], ev["row"], "OK", "-"])
            pos = idx + 1
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

    return rows, fail

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

def make_pdf(rows, cols, failure, fname, meta):
    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(A4),
        leftMargin=18,
        rightMargin=18,
        topMargin=18,
        bottomMargin=18
    )

    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph("<b>UFSBI Expert Failure Analysis Report V2</b>", styles["Title"]))
    story.append(Spacer(1, 8))
    story.append(Paragraph(f"<b>Generated:</b> {datetime.now().strftime('%d-%m-%Y %H:%M:%S')}", styles["Normal"]))
    story.append(Paragraph(f"<b>File:</b> {fname}", styles["Normal"]))

    if failure:
        story.append(Paragraph(f"<b>Failure Type:</b> {failure.get('failure_type', '-')}", styles["Normal"]))
        story.append(Paragraph(f"<b>Department:</b> {failure.get('department', '-')}", styles["Normal"]))
        if "reason" in failure:
            story.append(Paragraph(f"<b>Reason:</b> {failure.get('reason', '-')}", styles["Normal"]))

    story.append(Paragraph(f"<b>Detected Columns:</b> relay={meta['relay']}, status={meta['status']}, time={meta['time']}", styles["Normal"]))
    story.append(Spacer(1, 8))

    data = [cols] + [[str(x) for x in r] for r in rows]
    table = Table(data, repeatRows=1, colWidths=[30, 95, 55, 45, 115, 40, 55, 360])
    table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.35, colors.black),
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("FONTSIZE", (0, 0), (-1, -1), 6.5),
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

st.title("UFSBI Expert Analyzer V2")
st.write("Faulty UFSBI Data Logger Excel upload करें. यह Link Error और Signal Relay Sequence दोनों analyze करेगा.")

file = st.file_uploader("Upload Faulty Excel", type=["xls", "xlsx"])

if file:
    try:
        df = read_excel(file)
        events, meta = extract_events(df)

        cols = ["Step", "Phase", "Relay", "Expected", "Time", "Row", "Result", "Maintainer Check"]

        link_fail = detect_link_failure(events)

        if link_fail:
            rows = [[
                0,
                "Communication / Link",
                "BIPR/FR/TGT",
                "CHECK",
                "-",
                "-",
                "LINK ERROR",
                link_fail["action"]
            ]]
            fail = link_fail

            st.error(f"{link_fail['failure_type']} | Department: TELECOM")
            st.write(link_fail["reason"])
            st.write(link_fail["action"])

        else:
            rows, fail = analyze_signal_sequence(events)

            if fail:
                st.error(f"{fail['failure_type']}: Step {fail['step']} missing - {fail['relay']} {fail['status']}")
                st.write(fail["check"])
            else:
                st.success("All expert-defined relay events found. No sequence mismatch detected.")

        st.subheader("Analysis Result")
        st.markdown(html(rows, cols), unsafe_allow_html=True)

        pdf = make_pdf(rows, cols, fail, file.name, meta)

        st.download_button(
            "Download PDF Report",
            pdf,
            "UFSBI_Expert_Failure_Report_V2.pdf",
            "application/pdf"
        )

    except Exception as e:
        st.error(str(e))

else:
    st.info("Faulty UFSBI Excel upload करें.")
