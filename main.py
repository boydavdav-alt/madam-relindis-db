import warnings
warnings.filterwarnings("ignore")

import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from datetime import datetime
from PIL import Image
import os
import re
import io
import base64
import glob


st.set_page_config(page_title="CIMIC TECHNICAL INTELLIGENCE SYSTEM", layout="wide")


# ------------------------------
# GLOBALS
# ------------------------------
SAVE_FOLDER = "saved_reports"
LOGO_PATH = r"C:\Users\ivo\Desktop\STOP\dondt\assets\logo.png"
VIDEO_PATH = "radar_bg.mp4" # put your radar video here
COUNTER_FILE = "report_counter.txt"
EXCLUDE_NUMBERS = ["8445", "431", "8160"] # system codes to block

os.makedirs(SAVE_FOLDER, exist_ok=True)


# ------------------------------
# CSS + VIDEO BACKGROUND + GLASS DASHBOARD
# ------------------------------
def get_base64_video(video_path):
    if os.path.exists(video_path):
        with open(video_path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    return None


video_b64 = get_base64_video(VIDEO_PATH)
logo_b64 = ""
if os.path.exists(LOGO_PATH):
    with open(LOGO_PATH, "rb") as f: logo_b64 = base64.b64encode(f.read()).decode()


video_html = f'<video autoplay muted loop class="video-bg"><source src="data:video/mp4;base64,{video_b64}" type="video/mp4"></video>' if video_b64 else ""


st.markdown(f"""
{video_html}
<style>
@import url('https://fonts.googleapis.com/css2?family=Orbitron:wght@400;700;900&display=swap');
.video-bg {{position: fixed; top: 0; left: 0; width: 100vw; height: 100vh; object-fit: cover; z-index: -999; filter: brightness(0.35) contrast(1.3);}}
html, body, [data-testid="stAppViewContainer"],.main,.block-container {{background: transparent!important;}}
[data-testid="stHeader"], [data-testid="stToolbar"] {{background: rgba(0,0,0,0)!important;}}
.logo {{position: absolute; top: 15px; right: 20px; width: 80px; height: 80px; object-fit: contain; filter: drop-shadow(0 0 10px #00ddff); z-index: 999;}}
.container {{width: 98%; max-width: 1400px; margin: 10px auto; margin-top: 100px; background: rgba(0,8,25,0.55); border-radius: 12px; padding: 20px 25px; border: 1px solid rgba(0,200,255,0.5); box-shadow: 0 0 40px rgba(0,200,255,0.3); backdrop-filter: blur(15px);}}
.title {{font-family: 'Orbitron', sans-serif; font-size:28px; font-weight:900; color:#00ddff; text-align:center; letter-spacing: 3px; margin-bottom: 20px; text-shadow: 0 0 20px #00ddff; animation: glow 2s ease-in-out infinite alternate;}}
@keyframes glow {{from {{text-shadow: 0 0 10px #00ddff;}} to {{text-shadow: 0 0 25px #00ddff, 0 0 50px #00ddff;}}}}
.stButton>button {{width:100%; border-radius:8px; font-weight:bold; font-size:14px; background: rgba(0, 120, 255, 0.5); color:#00ddff; border:2px solid #00ddff; height:42px; backdrop-filter: blur(10px);}}
.stButton>button:hover {{background: rgba(0, 200, 255, 0.7); box-shadow: 0 0 25px #00ddff; transform: scale(1.05); color: white;}}
[data-testid="stFileUploader"], [data-testid="stMarkdownContainer"] {{background: rgba(0,8,25,0.5); border: 1px solid rgba(0,200,255,0.4); border-radius: 8px; padding: 10px; backdrop-filter: blur(10px); color: white!important;}}
label, p, div, span {{color: white!important;}}
h3 {{color: #00ddff; font-size: 18px; text-shadow: 0 0 10px #00ddff;}}
</style>
""", unsafe_allow_html=True)

if logo_b64:
    st.markdown(f'<img src="data:image/png;base64,{logo_b64}" class="logo">', unsafe_allow_html=True)

st.markdown('<div class="container">', unsafe_allow_html=True)
st.markdown('<div class="title">CIMIC TECHNICAL INTELLIGENCE SYSTEM</div>', unsafe_allow_html=True)


# ------------------------------
# REPORT NUMBER AUTO
# ------------------------------
def get_next_report_number():
    year = datetime.now().strftime("%Y"); n = 0
    if os.path.exists(COUNTER_FILE):
        try:
            with open(COUNTER_FILE, "r") as f: content = f.read().strip(); n = int(content.split("|")[-1]) if "|" in content else int(content)
        except: n = 0
    n += 1
    with open(COUNTER_FILE, "w") as f: f.write(f"{year}|{n}")
    return f"{n:03d}{year}"


# ------------------------------
# NORMALIZE + VALIDATE PHONE
# ------------------------------
def normalize_number(num):
    num = str(num).strip()
    if num.startswith("+"): return "+" + re.sub(r"\D", "", num)
    num = re.sub(r"\D", "", num)
    if num.startswith("237"): num = num[3:]
    return num


def is_valid_number(num):
    num = str(num).strip()
    if num in EXCLUDE_NUMBERS: return False
    if num.startswith("+"): return len(re.sub(r"\D", "", num)) >= 8
    if re.match(r'^6\d{8}$', num): return True # Only Cameroon 6XXXXXXXX
    return False


# ------------------------------
# BUILD NAME MAP FROM "Frequent Correspondent" D:E
# D = Telephone, E = Identite
# ------------------------------
def build_name_map(xl):
    name_map = {}
    sheets = xl.sheet_names
    if "Frequent Correspondent" in sheets:
        try:
            freq = xl.parse("Frequent Correspondent")
            # Col D = index 3, Col E = index 4
            for _, row in freq.iterrows():
                num = normalize_number(row.iloc[3]) if len(row) > 3 else ""
                name = str(row.iloc[4]).strip() if len(row) > 4 else ""
                if num and name and name!= "nan" and name!= "None":
                    name_map[num] = name
        except Exception as e:
            print("Error reading Frequent Correspondent:", e)
    return name_map


def format_contact(num, name_map):
    name = name_map.get(num, "").strip()
    return f"{num} - {name}" if name else num


# ------------------------------
# REGIONS
# ------------------------------
regions_cameroon = ["ADAMAWA","CENTER","EAST","FAR-NORTH","LITTORAL","NORTH","NORTH-WEST","SOUTH","SOUTH-WEST"]
def extract_region(text):
    text = str(text).upper()
    sorted_regions = sorted(regions_cameroon, key=len, reverse=True)
    for r in sorted_regions:
        if re.search(rf"\b{r}\b", text): return r
    return None


# ------------------------------
# ANALYZE FILE - YOUR ORIGINAL LOGIC + NAME LOOKUP
# ------------------------------
def analyze_file(file):
    data = {"number":"", "name":"", "operator":"","callers":[], "outgoing":[], "imei":[],"locations_full":[], "regions":[],"pattern":"", "device_usage":""}
    xl = pd.ExcelFile(file); sheets = xl.sheet_names

    name_map = build_name_map(xl) # NEW: get names from Frequent Correspondent

    try:
        sub = xl.parse("Abonné"); data["number"] = normalize_number(sub.iloc[0,0]); data["name"] = str(sub.iloc[0,1]); data["operator"] = str(sub.iloc[0,2])
    except: pass

    calls = sms = None
    if "Listing Appel" in sheets: calls = xl.parse("Listing Appel")
    if "Listing SMS" in sheets: sms = xl.parse("Listing SMS")
    if "Listing" in sheets and calls is None and sms is None: combined = xl.parse("Listing"); calls = sms = combined

    incoming, outgoing = [], []
    try:
        if calls is not None: r = calls.iloc[:,5].apply(normalize_number); f = calls[r == data["number"]]; incoming += f.iloc[:,0].apply(normalize_number).tolist()
        if sms is not None: r = sms.iloc[:,5].apply(normalize_number); f = sms[r == data["number"]]; incoming += f.iloc[:,0].apply(normalize_number).tolist()
    except: pass
    try:
        if calls is not None: r = calls.iloc[:,0].apply(normalize_number); f = calls[r == data["number"]]; outgoing += f.iloc[:,5].apply(normalize_number).tolist()
        if sms is not None: r = sms.iloc[:,0].apply(normalize_number); f = sms[r == data["number"]]; outgoing += f.iloc[:,5].apply(normalize_number).tolist()
    except: pass

    # FILTER + ADD NAMES
    incoming = [n for n in incoming if is_valid_number(n)]
    outgoing = [n for n in outgoing if is_valid_number(n)]

    data["callers"] = [format_contact(n, name_map) for n in pd.Series(incoming).value_counts().head(5).index.tolist()] if incoming else []
    data["outgoing"] = [format_contact(n, name_map) for n in pd.Series(outgoing).value_counts().head(5).index.tolist()] if outgoing else []

    locations, regions = [], []
    try:
        if calls is not None: f = calls[calls.iloc[:,0].apply(normalize_number) == data["number"]]
        seen = set()
        for v in f.iloc[:,1]:
            txt = str(v); norm = re.split(r'\(Cell:|\(|Long:|\)|Lat:|Azimut:', txt)[0].strip()
            if norm not in seen: locations.append(txt); seen.add(norm); r = extract_region(txt)
            if r: regions.append(r)
    except: pass
    data["locations_full"] = list(dict.fromkeys(locations)); data["regions"] = list(dict.fromkeys(regions))

    imei = []
    try:
        if calls is not None: f = calls[calls.iloc[:,0].apply(normalize_number) == data["number"]]; imei += f.iloc[:,2].dropna().astype(str).unique().tolist()
    except: pass
    imei = list(dict.fromkeys(imei))
    data["imei"] = imei
    data["device_usage"] = "Mainly one IMEI detected" if len(imei)<=1 else f"{len(imei)} different devices detected"
    data["pattern"] = "Mostly outgoing calls" if len(outgoing) > len(incoming) else "Mostly incoming calls"
    return data


# ------------------------------
# GENERATE REPORT - YOUR ORIGINAL FORMAT
# ------------------------------
def generate_report(analysis_results):
    doc = Document(); style = doc.styles['Normal']; style.font.name = "Times New Roman"; style.font.size = Pt(10); style.paragraph_format.space_after = Pt(0)
    section = doc.sections[0]; section.orientation = WD_ORIENT.LANDSCAPE; section.page_width, section.page_height = section.page_height, section.page_width
    section.left_margin = Cm(0.2); section.right_margin = Cm(0.2); section.top_margin = Cm(0.2); section.bottom_margin = Cm(0.2)

    header = section.header; header_para = header.paragraphs[0]; header_para.clear(); header_para.alignment = WD_ALIGN_PARAGRAPH.LEFT; header_para.paragraph_format.left_indent = Cm(0.2)
    if os.path.exists(LOGO_PATH): run = header_para.add_run(); run.add_picture(LOGO_PATH, width=Cm(1.5))

    footer = section.footer; p = footer.paragraphs[0]; p.text = "- Defence Confidential -"; p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p = doc.add_paragraph(); r = p.add_run("TECHNICAL ANALYSIS REPORT"); r.font.size = Pt(18); r.font.bold = True; r.font.underline = True; r.font.color.rgb = RGBColor(0,0,139); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    text_left_indent = Cm(1); today = datetime.now().strftime("%d %B %Y")
    p = doc.add_paragraph(); p.paragraph_format.left_indent = text_left_indent; p.add_run("CIMIC CENTER LE: "); r = p.add_run(today); r.font.color.rgb = RGBColor(255,0,0)
    report = get_next_report_number()
    p = doc.add_paragraph(); p.paragraph_format.left_indent = text_left_indent; p.add_run("N° "); r = p.add_run(report); r.font.color.rgb = RGBColor(255,0,0)

    numbers = [d["number"] for d in analysis_results]; p = doc.add_paragraph(); p.paragraph_format.left_indent = text_left_indent
    intro_text = "After reviewing the number(s) provided, the individual(s) registered their SIM card with the following listed below:" if len(numbers) == 1 else "After reviewing the provided numbers, the individuals registered their SIM cards with the following listed below:"
    p.add_run("Number(s) under investigation: ")
    for i, n in enumerate(numbers):
        r = p.add_run(n); r.font.color.rgb = RGBColor(0,128,0)
        if i < len(numbers)-1: p.add_run(", ")

    p = doc.add_paragraph(); p.paragraph_format.left_indent = text_left_indent; r = p.add_run(intro_text); r.font.color.rgb = RGBColor(128,0,128)

    for idx, d in enumerate(analysis_results, start=1):
        p = doc.add_paragraph(f"{idx}. SUBSCRIBER IDENTITY:"); p.runs[0].font.bold = True; p.runs[0].font.color.rgb = RGBColor(0,0,255); p.paragraph_format.left_indent = text_left_indent
        for field in ["number","name","operator"]: p = doc.add_paragraph(f"{field.capitalize()}: {d[field]}"); p.paragraph_format.left_indent = text_left_indent

        # NOW SHOWS NAME NEXT TO NUMBER
        p = doc.add_paragraph(f"{idx}.1 People who call the Subscriber Frequently:"); p.runs[0].font.bold = True; p.runs[0].font.color.rgb = RGBColor(0,0,255); p.paragraph_format.left_indent = text_left_indent
        for v in d["callers"]: p = doc.add_paragraph(f"• {v}"); p.paragraph_format.left_indent = text_left_indent

        p = doc.add_paragraph(f"{idx}.2 People He Calls Frequently:"); p.runs[0].font.bold = True; p.runs[0].font.color.rgb = RGBColor(0,0,255); p.paragraph_format.left_indent = text_left_indent
        for v in d["outgoing"]: p = doc.add_paragraph(f"• {v}"); p.paragraph_format.left_indent = text_left_indent

        p = doc.add_paragraph(f"{idx}.3 Locations detected of the Subscriber:"); p.runs[0].font.bold = True; p.runs[0].font.color.rgb = RGBColor(0,0,255); p.paragraph_format.left_indent = text_left_indent
        for v in d["locations_full"]: p = doc.add_paragraph(f"• {v}"); p.paragraph_format.left_indent = text_left_indent

        p = doc.add_paragraph(f"{idx}.4 Phone Device (IMEI):"); p.runs[0].font.bold = True; p.runs[0].font.color.rgb = RGBColor(0,0,255); p.paragraph_format.left_indent = text_left_indent
        for v in d["imei"]: p = doc.add_paragraph(f"• {v}"); p.paragraph_format.left_indent = text_left_indent

        p = doc.add_paragraph(f"{idx}.✓ Key Intelligence Summary"); p.runs[0].font.bold = True; p.runs[0].font.color.rgb = RGBColor(0,0,255); p.paragraph_format.left_indent = text_left_indent
        summary = [f"- Owner: {d['name']}",f"- Operator: {d['operator']}",f"- Main contacts: {', '.join(d['outgoing'])}",f"- Frequent incoming contacts: {', '.join(d['callers'])}",f"- Device usage: {d['device_usage']}",f"- Communication pattern: {d['pattern']}",f"- Location data: {', '.join(d['regions'])}"]
        for s in summary: p = doc.add_paragraph(s); p.runs[0].font.color.rgb = RGBColor(255,0,0); p.paragraph_format.left_indent = text_left_indent
    doc.add_paragraph(""); p = doc.add_paragraph(); r = p.add_run("LE COM CIMIC/BIR"); r.font.bold = True; r.font.size = Pt(20); r.font.underline = True; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    buf = io.BytesIO(); doc.save(buf); buf.seek(0); return buf


# ------------------------------
# UI - YOUR ORIGINAL 3 BUTTONS
# ------------------------------
uploaded_files = st.file_uploader("📁 Upload XLSX Files", type=["xlsx"], accept_multiple_files=True)
col1, col2, col3 = st.columns(3)
with col1: analyze_btn = st.button("⚡ Analyze Automatically")
with col2: report_btn = st.button("📄 Generate Report")
with col3: clear_btn = st.button("🗑️ Clear")

if 'results' not in st.session_state: st.session_state['results'] = []
if clear_btn: st.session_state['results'] = []; st.rerun()

if analyze_btn:
    if not uploaded_files: st.warning("Please upload files first")
    else:
        with st.spinner("Analyzing..."):
            results = []
            for f in uploaded_files: results.append(analyze_file(f))
            st.session_state['results'] = results
            st.success(f"{len(results)} files analyzed")

if st.session_state['results']:
    st.subheader("Analysis Results")
    for r in st.session_state['results']:
        st.markdown(f"**Number:** `{r['number']}` | **Name:** {r['name']} | **Operator:** {r['operator']}")
        st.markdown(f"**Outgoing with Names:** {r['outgoing']}")
        st.markdown(f"**Incoming with Names:** {r['callers']}")
        st.divider()

if report_btn:
    if not st.session_state['results']: st.warning("Run analysis first")
    else:
        doc_buffer = generate_report(st.session_state['results'])
        st.download_button(label="📥 Download Word Report", data=doc_buffer, file_name=f"Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx")

st.markdown('</div>', unsafe_allow_html=True)