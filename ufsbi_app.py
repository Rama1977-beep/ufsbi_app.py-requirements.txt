import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

st.set_page_config(page_title="UFSBI Expert Analyzer", layout="wide")

EXPERT_SEQUENCE = [
    [1,"Line Clear Initiating","BPNR","UP","BPNR not pickup: check bell button contact and SM key."],
    [2,"Line Clear Initiating","TGTNR","UP","TGTNR not pickup: check TGT button and CNR relay contact D7/D8."],
    [3,"Line Clear Initiating","TGTXR","UP","TGTXR not pickup: check TGTR drop A7/A8, BPNR B1/N2, TGTNR A1/A2, ASGNCR A1/A2, DAZTR B1/B2, 24V fuse."],
    [4,"Line Clear Link","TGTYR","UP","TGTYR not pickup: check R1/R2. Ask STN-B whether TCFR showing. If not, fault at STN-B. Check STN-A ASGNCR B1/N2."],
    [5,"Train On Line Prep","HSGNCR","DOWN","HSGNCR not drop: check HSGNCR drop path."],
    [6,"Train On Line Prep","HSATPR","DOWN","HSATPR not drop: check HSATPR drop proving."],
    [7,"Train On Line Prep","TAR1","UP","TAR1 not pickup/hold: check HSATPR A7/A8, TAR1 A2/A1, HSBTPR D2/D1, TCFR A3/A4, TAR2 D5/D6."],
    [8,"Train On Line Prep","HSBTPR","DOWN","HSBTPR not drop: check HSBTPR drop path."],
    [9,"Train On Line Prep","HSATPR","UP","HSATPR not pickup: check BTSR A5/A6, HSBTPR D7/D8, HSATPR A1/A2, TAR1 B1/N2, TAR2 D2/D1."],
    [10,"Train On Line Prep","TAR2","UP","TAR2 not pickup/hold: check BTSR A5/A6, HSBTPR D7/D8, HSATPR A1/A2, TAR1 B1/N2, TAR2 D2/D1."],
    [11,"Train On Line Prep","TAR1","DOWN","TAR1 not drop: check TAR1 drop circuit/holding path."],
    [12,"Train On Line","DAZTR","UP","DAZTR not pickup: check DAZTR pickup circuit and 24V feed."],
    [13,"Train On Line","HSGNCR","UP","HSGNCR not pickup: check HSGNCR pickup circuit."],
    [14,"Train On Line","TCFR","DOWN","TCFR not drop: check TCFXR B7/B8, TAR2 C2/C1, CAR D8/D7, ASGNCPR A1/A2, CNR D6/D5, HSGNCR D1/D2, DAZTR C1/C2, LCB key reverse."],
    [15,"Train On Line","BTSR","UP","BTSR not pickup: check RAZTR C4/C5, CAR D5/D6, TCFR D7/D8, BTSR A2/A1."],
    [16,"Train On Line Link","TGTNZR","DOWN","TGTNZR not drop: check STN-B."],
    [17,"Train On Line","TGTR","UP","TGTR not pickup: check TGTR R1/R2 voltage and previous relay contacts."],
    [18,"Proving","DAZTR","UP","Check DAZTR contact B1/N2."],
    [19,"Proving","ASGNCR","UP","Check ASGNCR contact A1/A2."],
    [20,"Proving","BPNR","UP","Check BPNR contact B1/N2."],
    [21,"Proving","TGTYR","UP","Check TGTYR contact A1/A2."],
    [22,"Proving","TGTXR","UP","Check TGTXR contact A1/D2."]
]

def s(x):
    return "" if pd.isna(x) else str(x)

def clean_relay(x):
    t = s(x).upper().strip()
    t = re.sub(r"\(.*?\)", "", t)
    return re.sub(r"\s+", "", t)

def clean_status(x):
    t = s(x).upper().strip()
    if t in ["PICKUP","PICK UP","PICK","PU","ON"]:
        return "UP"
    if t in ["DROP","DROPPED","OFF"]:
        return "DOWN"
    return t

def read_excel(file):
    for eng in ["openpyxl","xlrd",None]:
        try:
            file.seek(0)
            return pd.read_excel(file, engine=eng, header=None, dtype=str) if eng else pd.read_excel(file, header=None, dtype=str)
        except Exception:
            pass
    raise Exception("Excel file read नहीं हो रहा.")

def detect_columns(df):
    relay_col = status_col = time_col = None
    best_relay = best_status = best_time = -1
    for c in df.columns:
        vals = df[c].fillna("").astype(str).head(1000)
        r = sum(5 if "(" in v and ")" in v else 1 if re.search(r"[A-Z]{2,}", v.upper()) else 0 for v in vals)
        st = sum(clean_status(v) in ["UP","DOWN"] for v in vals)
        tm = sum(bool(re.search(r"\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4}.*\d{1,2}:\d{2}", v)) for v in vals)
        if r > best_relay: best_relay, relay_col = r, c
        if st > best_status: best_status, status_col = st, c
        if tm > best_time: best_time, time_col = tm, c
    return relay_col, status_col, time_col

def extract_events(df):
    rc, sc, tc = detect_columns(df)
    events = []
    for i,row in df.iterrows():
        relay = clean_relay(row[rc])
        status = clean_status(row[sc])
        if relay and status in ["UP","DOWN"]:
            events.append({"row":i+1,"relay":relay,"status":status,"time":s(row[tc]) if tc is not None else "-"})
    return events, {"relay":rc,"status":sc,"time":tc}

def find_event(events, relay, status, start):
    for i in range(start, len(events)):
        if events[i]["relay"] == relay and events[i]["status"] == status:
            return i, events[i]
    return None, None

def classify(step):
    if step <= 4: return "Line Clear Not Initiating / Link Failure"
    if step <= 11: return "Train On Line Initiation Not Completed"
    if step <= 17: return "Train On Line Not Completed"
    return "Arrival / Proving Not Completed"

def analyze(events):
    rows, pos, fail = [], 0, None
    for step, phase, relay, status, check in EXPERT_SEQUENCE:
        idx, ev = find_event(events, relay, status, pos)
        if ev:
            rows.append([step,phase,relay,status,ev["time"],ev["row"],"OK","-"])
            pos = idx + 1
        else:
            rows.append([step,phase,relay,status,"-","-","MISSING",check])
            fail = {"step":step,"phase":phase,"relay":relay,"status":status,"check":check}
            break
    return rows, fail

def html(rows, cols):
    h="<table style='border-collapse:collapse;width:100%;font-size:13px'><tr>"
    h += "".join(f"<th style='border:1px solid #aaa;padding:6px;background:#eee'>{c}</th>" for c in cols)+"</tr>"
    for r in rows:
        h += "<tr>"+"".join(f"<td style='border:1px solid #aaa;padding:6px;vertical-align:top'>{x}</td>" for x in r)+"</tr>"
    return h+"</table>"

def make_pdf(rows, cols, fail, fname, meta):
    buf=BytesIO()
    doc=SimpleDocTemplate(buf,pagesize=landscape(A4),leftMargin=18,rightMargin=18,topMargin=18,bottomMargin=18)
    styles=getSampleStyleSheet()
    story=[Paragraph("<b>UFSBI Expert Failure Analysis Report</b>",styles["Title"]),Spacer(1,8)]
    story.append(Paragraph(f"<b>Generated:</b> {datetime.now().strftime('%d-%m-%Y %H:%M:%S')}",styles["Normal"]))
    story.append(Paragraph(f"<b>File:</b> {fname}",styles["Normal"]))
    story.append(Paragraph(f"<b>Failure Type:</b> {classify(fail['step']) if fail else 'No mismatch found'}",styles["Normal"]))
    story.append(Paragraph(f"<b>Columns:</b> relay={meta['relay']}, status={meta['status']}, time={meta['time']}",styles["Normal"]))
    story.append(Spacer(1,8))
    data=[cols]+[[str(x) for x in r] for r in rows]
    table=Table(data,repeatRows=1,colWidths=[30,95,55,45,115,40,55,360])
    table.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.35,colors.black),("BACKGROUND",(0,0),(-1,0),colors.lightgrey),("FONTSIZE",(0,0),(-1,-1),6.5),("VALIGN",(0,0),(-1,-1),"TOP")]))
    story.append(table)
    if fail:
        story.append(Spacer(1,10))
        story.append(Paragraph("<b>Failure Pinpointed</b>",styles["Heading2"]))
        story.append(Paragraph(f"Step {fail['step']}: {fail['relay']} {fail['status']} missing",styles["Normal"]))
        story.append(Paragraph(f"<b>Maintainer Action:</b> {fail['check']}",styles["Normal"]))
    doc.build(story)
    buf.seek(0)
    return buf

st.title("UFSBI Expert Analyzer")
st.write("Faulty UFSBI Data Logger Excel upload करें. Healthy file की जरूरत नहीं है.")

file = st.file_uploader("Upload Faulty Excel", type=["xls","xlsx"])

if file:
    try:
        df = read_excel(file)
        events, meta = extract_events(df)
        rows, fail = analyze(events)
        cols=["Step","Phase","Relay","Expected","Time","Row","Result","Maintainer Check"]
        st.markdown(html(rows, cols), unsafe_allow_html=True)
        if fail:
            st.error(f"{classify(fail['step'])}: Step {fail['step']} missing - {fail['relay']} {fail['status']}")
            st.write(fail["check"])
        else:
            st.success("All expert-defined relay events found.")
        pdf = make_pdf(rows, cols, fail, file.name, meta)
        st.download_button("Download PDF Report", pdf, "UFSBI_Expert_Failure_Report.pdf", "application/pdf")
    except Exception as e:
        st.error(str(e))
else:
    st.info("Faulty UFSBI Excel upload करें.")
