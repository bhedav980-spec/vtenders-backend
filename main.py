from fastapi import FastAPI, UploadFile, File
from fastapi.responses import HTMLResponse, FileResponse
from typing import List
import fitz
import pandas as pd
import uuid
import os
import re
import json

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

try:
    from openai import OpenAI
except Exception:
    OpenAI = None

app = FastAPI()

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
client = OpenAI(api_key=OPENAI_API_KEY) if OpenAI and OPENAI_API_KEY else None


# ---------------- KNOWLEDGE BASE ----------------
WORK_KNOWLEDGE_BASE = {
    "cable": {
        "keywords": ["cable", "xlpe", "ht cable", "lt cable", "cable laying", "cable pulling", "termination"],
        "work_type": "Cable Laying / Termination Work",
        "material": "Cable, lugs, glands, ferrules, tags, cable ties, saddles, warning tape if specified",
        "method": "Check route, handle cable drum, place rollers, pull cable safely, dress cable, gland/terminate, test insulation and continuity.",
        "tools": "Cable roller, drum jack, winch, crimping tool, megger, multimeter, PPE",
        "labour": "Cable gang 4-6 persons + electrician + supervisor",
        "time": "Depends on cable size and route; approx. 1 day per 300-500 meter for normal route"
    },
    "earthing": {
        "keywords": ["earthing", "earth pit", "earth electrode", "gi strip", "copper strip", "chemical earthing"],
        "work_type": "Earthing Work",
        "material": "Earth electrode, GI/Copper strip, earth compound, chamber cover, nuts/bolts",
        "method": "Excavate pit, install electrode, connect strip, backfill with compound, water and test earth resistance.",
        "tools": "Earth tester, spade, welding/drilling tools, multimeter",
        "labour": "Electrician + 2 labour + supervisor",
        "time": "Approx. 2-4 hours per earth pit depending on soil/site condition"
    },
    "concrete": {
        "keywords": ["concrete", "rcc", "pcc", "m20", "m25", "m30", "cement concrete"],
        "work_type": "Concrete Work",
        "material": "Cement, sand, aggregate, water, admixture if specified, reinforcement steel if RCC",
        "method": "Batching, mixing, placing, vibration/compaction, levelling, finishing and curing as per specification.",
        "tools": "Mixer, vibrator, cube mould, trowel, level, curing arrangement",
        "labour": "Mason + 4-6 labour + supervisor",
        "time": "Depends on quantity; approx. 8-12 Cum/day with small team and mixer"
    },
    "excavation": {
        "keywords": ["excavation", "earthwork", "earth work", "trench", "pit", "digging", "backfilling"],
        "work_type": "Excavation / Earthwork",
        "material": "No permanent material; soil disposal/backfilling material as applicable",
        "method": "Marking, excavation manually/by machine, dressing, dewatering if required, disposal/backfilling.",
        "tools": "JCB, spade, pickaxe, dumper, measuring tape",
        "labour": "Machine operator or 4-6 labour + supervisor",
        "time": "Depends on depth and soil; machine excavation faster than manual"
    },
    "demolition": {
        "keywords": ["demolition", "dismantling", "dismantle", "breaking", "removal", "dispose", "disposal"],
        "work_type": "Demolition / Dismantling Work",
        "material": "No permanent material required; debris handling and consumables only",
        "method": "Inspect site, isolate utilities, barricade area, dismantle/break carefully, stack serviceable material and dispose debris.",
        "tools": "Breaker, hammer, chisel, cutter, wheelbarrow, barricade, PPE",
        "labour": "Demolition labour 4-6 persons + supervisor",
        "time": "Depends on structure and safety restrictions"
    },
    "panel": {
        "keywords": ["panel", "switchgear", "mccb", "acb", "lt panel", "control panel", "distribution board"],
        "work_type": "Electrical Panel Work",
        "material": "Panel, breaker, busbar, control wiring, lugs, glands, ferrules and accessories",
        "method": "Position panel, align/fix, cable termination, control wiring, testing and commissioning.",
        "tools": "Crimping tool, multimeter, torque wrench, megger",
        "labour": "Electrician + technician + supervisor",
        "time": "Approx. 1-2 days per panel depending on size and termination"
    },
    "scada": {
        "keywords": ["scada", "rtu", "plc", "hmi", "automation", "remote monitoring", "data logger"],
        "work_type": "SCADA / Automation Work",
        "material": "PLC/RTU, HMI, sensors, communication module, control cable, panel accessories",
        "method": "Install devices, complete wiring, configure communication, test signals and commission system.",
        "tools": "Laptop, multimeter, crimping tool, communication tester",
        "labour": "Automation engineer + technician",
        "time": "Approx. 1-3 days depending on points and communication complexity"
    },
    "painting": {
        "keywords": ["painting", "paint", "primer", "coating", "enamel", "epoxy"],
        "work_type": "Painting / Coating Work",
        "material": "Primer, paint/coating, thinner, putty if specified",
        "method": "Surface preparation, primer coat, paint/coating application and finishing.",
        "tools": "Brush, roller, spray gun, scraper, sandpaper",
        "labour": "Painter + helper",
        "time": "Depends on area and number of coats"
    },
    "masonry": {
        "keywords": ["brick", "brickwork", "block work", "masonry", "aac block", "stone masonry"],
        "work_type": "Masonry Work",
        "material": "Bricks/blocks/stone, cement mortar, sand and water",
        "method": "Line-level setting, mortar preparation, laying, joint filling and curing.",
        "tools": "Trowel, line dori, level tube, plumb bob, scaffolding",
        "labour": "Mason + helper",
        "time": "Depends on wall thickness and height"
    },
    "plaster": {
        "keywords": ["plaster", "plastering", "cement plaster"],
        "work_type": "Plaster Work",
        "material": "Cement, sand, water and curing arrangement",
        "method": "Prepare surface, apply mortar, level, finish and cure.",
        "tools": "Trowel, level, wooden float, scaffolding",
        "labour": "Mason + helper",
        "time": "Depends on area and thickness"
    },
    "solar": {
        "keywords": ["solar", "module", "pv", "inverter", "string", "mms", "dc cable", "ac cable"],
        "work_type": "Solar Plant Work",
        "material": "PV module, MMS, inverter, cable, connector, earthing and accessories as specified",
        "method": "Install structure/modules, do cabling, termination, testing and commissioning.",
        "tools": "Torque wrench, multimeter, megger, MC4 crimper, PPE",
        "labour": "Solar technician + electrician + supervisor",
        "time": "Depends on plant size and manpower"
    }
}


# ---------------- PDF TEXT EXTRACTION ----------------
def extract_text_from_pdf(file_bytes):
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    lines = []
    pages_text = []

    for page_no, page in enumerate(doc, start=1):
        text = page.get_text("text")
        page_lines = []

        for line in text.splitlines():
            line = re.sub(r"\s+", " ", line).strip()
            if line:
                lines.append({"page": page_no, "text": line})
                page_lines.append(line)

        pages_text.append({
            "page": page_no,
            "text": "\n".join(page_lines)
        })

    return lines, pages_text


# ---------------- SECTION DETECTION ----------------
def detect_section(line):
    t = line.lower()

    if any(x in t for x in ["bill of quantity", "boq", "schedule of quantity", "price schedule"]):
        return "BOQ"
    if any(x in t for x in ["scope of work", "scope"]):
        return "SCOPE_OF_WORK"
    if any(x in t for x in ["technical specification", "specification", "specifications"]):
        return "TECHNICAL_SPECIFICATION"
    if any(x in t for x in ["approved make", "make list", "list of approved"]):
        return "APPROVED_MAKE"
    if any(x in t for x in ["vendor list", "approved vendor"]):
        return "VENDOR_LIST"
    if any(x in t for x in ["general condition", "gcc"]):
        return "GENERAL_CONDITION"
    if any(x in t for x in ["special condition", "scc"]):
        return "SPECIAL_CONDITION"
    if any(x in t for x in ["safety", "ppe", "hse"]):
        return "SAFETY"
    if any(x in t for x in ["drawing", "drawing no", "drawing number"]):
        return "DRAWING"

    return "OTHER"


def build_context(lines):
    current_section = "OTHER"
    context_rows = []

    for row in lines:
        section = detect_section(row["text"])
        if section != "OTHER":
            current_section = section

        context_rows.append({
            "page": row["page"],
            "section": current_section,
            "text": row["text"]
        })

    return context_rows


# ---------------- TEXT CLEANING ----------------
def merge_lines_for_boq(context_rows):
    merged = []
    buffer = ""
    page = None
    section = "OTHER"

    item_start = re.compile(r"^\s*(\d+|[0-9]+\.[0-9]+)\s+")

    for row in context_rows:
        line = row["text"]

        if item_start.match(line):
            if buffer:
                merged.append({"page": page, "section": section, "text": buffer.strip()})
            buffer = line
            page = row["page"]
            section = row["section"]
        else:
            if buffer:
                buffer += " " + line

    if buffer:
        merged.append({"page": page, "section": section, "text": buffer.strip()})

    return merged


# ---------------- BOQ EXTRACTION ----------------
def normalize_unit(unit):
    if not unit:
        return ""
    u = unit.strip().lower()

    unit_map = {
        "cum": "Cum", "cu.m": "Cum", "cu.m.": "Cum", "m3": "Cum", "m³": "Cum", "cmt": "Cum",
        "sqm": "Sqm", "sq.m": "Sqm", "sq.m.": "Sqm", "m2": "Sqm", "m²": "Sqm", "smt": "Sqm",
        "rmt": "Rmt", "rm": "Rmt", "mtr": "Mtr", "meter": "Mtr", "metre": "Mtr", "m": "Mtr",
        "nos": "Nos", "no": "Nos", "each": "Nos", "set": "Set", "job": "Job", "lot": "Lot",
        "kg": "Kg", "mt": "MT", "ton": "MT", "ltr": "Ltr", "litre": "Ltr"
    }

    return unit_map.get(u, unit)


def extract_boq(context_rows):
    boq = []
    merged_rows = merge_lines_for_boq(context_rows)

    known_units = r"(cum|cu\.m\.?|m3|m³|cmt|sqm|sq\.m\.?|m2|m²|smt|rmt|rm|mtr|meter|metre|m|nos|no|each|set|job|lot|kg|mt|ton|ltr|litre)"

    pattern = re.compile(
        rf"^\s*(?P<item>\d+(\.\d+)*)\s+"
        rf"(?P<desc>.*?)\s+"
        rf"(?P<qty>\d+(\.\d+)?)\s*"
        rf"(?P<unit>{known_units})\s+"
        rf"(?P<rate>\d+(\.\d+)?)",
        re.IGNORECASE
    )

    for row in merged_rows:
        line = row["text"]
        clean = line.replace(",", "").replace("₹", "")

        match = pattern.search(clean)

        if match:
            boq.append({
                "Item No": match.group("item"),
                "Description": match.group("desc").strip(),
                "Quantity": match.group("qty"),
                "Unit": normalize_unit(match.group("unit")),
                "Rate": match.group("rate"),
                "Page": row["page"],
                "Section": row["section"],
                "Source Line": line
            })

    return boq


# ---------------- RELATED CONTEXT FINDER ----------------
def get_relevant_context(description, context_rows, max_lines=18):
    desc_words = set(re.findall(r"[a-zA-Z]{4,}", description.lower()))
    scored = []

    important_sections = [
        "TECHNICAL_SPECIFICATION",
        "SCOPE_OF_WORK",
        "APPROVED_MAKE",
        "VENDOR_LIST",
        "SPECIAL_CONDITION",
        "GENERAL_CONDITION",
        "SAFETY",
        "DRAWING"
    ]

    for row in context_rows:
        text_lower = row["text"].lower()
        score = 0

        for w in desc_words:
            if w in text_lower:
                score += 2

        if row["section"] in important_sections:
            score += 1

        if score > 0:
            scored.append((score, row))

    scored.sort(reverse=True, key=lambda x: x[0])
    selected = scored[:max_lines]

    return "\n".join(
        [f"[Page {r['page']} | {r['section']}] {r['text']}" for _, r in selected]
    )


# ---------------- RULE FALLBACK ANALYSIS ----------------
def fallback_analyze(description, unit):
    desc = description.lower()
    unit_l = str(unit).lower()

    best_score = 0
    best = None

    for _, data in WORK_KNOWLEDGE_BASE.items():
        score = 0

        for kw in data["keywords"]:
            if kw in desc:
                score += 3

        if unit_l in ["cum"] and any(x in desc for x in ["concrete", "excavation", "earthwork", "pcc", "rcc"]):
            score += 2

        if unit_l in ["sqm"] and any(x in desc for x in ["painting", "plaster", "flooring", "tile"]):
            score += 2

        if unit_l in ["mtr", "rmt"] and any(x in desc for x in ["cable", "pipe", "strip"]):
            score += 2

        if score > best_score:
            best_score = score
            best = data

    if best:
        confidence = "High" if best_score >= 5 else "Medium"
        return {
            "content_type": "WORK_ITEM",
            "identified_work_type": best["work_type"],
            "related_conditions": "-",
            "technical_specification": "-",
            "applicable_standard": "-",
            "material_required": best["material"],
            "approved_make": "-",
            "vendor_list": "-",
            "execution_method": best["method"],
            "tools_equipment": best["tools"],
            "labour_requirement": best["labour"],
            "supervisor_requirement": "Supervisor required as per site condition",
            "estimated_time": best["time"],
            "risk_review_point": "-",
            "confidence": confidence
        }

    return {
        "content_type": "WORK_ITEM",
        "identified_work_type": "General Tender Work Item",
        "related_conditions": "-",
        "technical_specification": "-",
        "applicable_standard": "-",
        "material_required": "As per tender item description and specification",
        "approved_make": "-",
        "vendor_list": "-",
        "execution_method": "Execute as per tender specification, drawings, applicable standards and site engineer instruction",
        "tools_equipment": "Standard tools, tackles, measuring instruments and PPE as required",
        "labour_requirement": "Skilled / semi-skilled labour as required",
        "supervisor_requirement": "Site supervisor required",
        "estimated_time": "To be estimated based on quantity, site access and manpower",
        "risk_review_point": "Specific work type not strongly identified; verify with tender specification",
        "confidence": "Low"
    }


# ---------------- AI ANALYSIS ----------------
def safe_json_loads(text):
    try:
        text = text.strip()
        text = text.replace("```json", "").replace("```", "").strip()
        return json.loads(text)
    except Exception:
        return None


def ai_analyze_boq_item(item, relevant_context):
    fallback = fallback_analyze(item["Description"], item["Unit"])

    if client is None:
        return fallback

    prompt = f"""
You are a senior tender analysis engineer.

Analyze this BOQ/work item and related tender text. Identify whether it is a work item and map tender clauses/specifications.

Return ONLY valid JSON. No markdown. No explanation.

JSON keys:
content_type
identified_work_type
related_conditions
technical_specification
applicable_standard
material_required
approved_make
vendor_list
execution_method
tools_equipment
labour_requirement
supervisor_requirement
estimated_time
risk_review_point
confidence

Rules:
- Use tender context where available.
- If approved make/vendor/standard is not found, write "-".
- Do not hallucinate tender-specific clauses.
- Method/tools/labour/time may use engineering knowledge if tender is silent.
- confidence must be High, Medium, or Low.

BOQ ITEM:
Item No: {item["Item No"]}
Description: {item["Description"]}
Quantity: {item["Quantity"]}
Unit: {item["Unit"]}
Rate: {item["Rate"]}

RELATED TENDER CONTEXT:
{relevant_context}
"""

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are a tender analysis expert. Return only valid JSON."},
                {"role": "user", "content": prompt}
            ],
            temperature=0
        )

        content = response.choices[0].message.content
        data = safe_json_loads(content)

        if not data:
            return fallback

        final = fallback.copy()
        for k in final.keys():
            if k in data and str(data[k]).strip():
                final[k] = data[k]

        return final

    except Exception as e:
        fallback["risk_review_point"] = f"AI analysis failed, fallback used: {str(e)}"
        fallback["confidence"] = "Medium" if fallback["confidence"] == "High" else fallback["confidence"]
        return fallback


# ---------------- RATE ----------------
def get_rate(rate):
    try:
        return float(str(rate).replace(",", "").replace("₹", ""))
    except Exception:
        return 0


# ---------------- EXCEL FORMATTING ----------------
def format_excel(file_name):
    wb = load_workbook(file_name)
    ws = wb.active
    ws.title = "Vtenders AI Output"

    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    header_fill = PatternFill(start_color="0B3D91", end_color="0B3D91", fill_type="solid")
    ok_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    medium_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    low_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = thin_border

    ws.freeze_panes = "A2"

    confidence_col = None
    for cell in ws[1]:
        if cell.value == "Confidence":
            confidence_col = cell.column

    if confidence_col:
        for row in range(2, ws.max_row + 1):
            c = ws.cell(row=row, column=confidence_col)
            if c.value == "High":
                c.fill = ok_fill
            elif c.value == "Medium":
                c.fill = medium_fill
            else:
                c.fill = low_fill

    widths = {
        "A": 10, "B": 55, "C": 26, "D": 16, "E": 12, "F": 12,
        "G": 16, "H": 45, "I": 55, "J": 45, "K": 45,
        "L": 45, "M": 55, "N": 45, "O": 35, "P": 35,
        "Q": 35, "R": 45, "S": 14, "T": 70, "U": 10, "V": 18
    }

    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    for row in range(1, ws.max_row + 1):
        ws.row_dimensions[row].height = 55

    wb.save(file_name)


# ---------------- HOME ----------------
@app.get("/", response_class=HTMLResponse)
def home():
    ai_status = "Enabled" if client else "Disabled - OPENAI_API_KEY not found"
    return f"""
    <html>
    <body style="font-family: Arial; padding: 30px;">
        <h2>Vtenders AI Tender Analyzer</h2>
        <p>AI Status: <b>{ai_status}</b></p>
        <p>Upload multiple tender PDF files. System will analyze BOQ, specifications, conditions, materials, method, tools, manpower and time.</p>
        <form action="/upload/" method="post" enctype="multipart/form-data">
            <input type="file" name="files" multiple required>
            <br><br>
            <button type="submit">Upload & Analyze</button>
        </form>
    </body>
    </html>
    """


@app.get("/ai-test")
def ai_test():
    sample_item = {
        "Item No": "1",
        "Description": "Providing and laying 3.5 core 240 sqmm XLPE cable including termination",
        "Quantity": "100",
        "Unit": "Mtr",
        "Rate": "500"
    }

    result = ai_analyze_boq_item(sample_item, "No tender context available for test.")
    return {"ai_enabled": client is not None, "result": result}


# ---------------- MAIN UPLOAD ----------------
@app.post("/upload/")
async def upload_files(files: List[UploadFile] = File(...)):
    all_items = []

    for file in files:
        content = await file.read()

        lines, pages_text = extract_text_from_pdf(content)
        context_rows = build_context(lines)
        boq = extract_boq(context_rows)

        for item in boq:
            relevant_context = get_relevant_context(item["Description"], context_rows)
            analysis = ai_analyze_boq_item(item, relevant_context)

            qty = 0
            try:
                qty = float(str(item["Quantity"]).replace(",", ""))
            except Exception:
                pass

            rate = get_rate(item["Rate"])
            amount = qty * rate

            all_items.append({
                "Item No": item["Item No"],
                "Tender Work Item": item["Description"],
                "Identified Work Type": analysis["identified_work_type"],
                "Quantity": qty,
                "Unit": item["Unit"],
                "Rate": rate,
                "Amount": amount,
                "Content Type": analysis["content_type"],
                "Related Tender Conditions": analysis["related_conditions"],
                "Technical Specification": analysis["technical_specification"],
                "Applicable Standard": analysis["applicable_standard"],
                "Material Required": analysis["material_required"],
                "Approved Make": analysis["approved_make"],
                "Vendor List": analysis["vendor_list"],
                "Execution Method": analysis["execution_method"],
                "Tools & Equipment": analysis["tools_equipment"],
                "Labour Requirement": analysis["labour_requirement"],
                "Supervisor Requirement": analysis["supervisor_requirement"],
                "Estimated Time": analysis["estimated_time"],
                "Risk / Review Point": analysis["risk_review_point"],
                "Confidence": analysis["confidence"],
                "Source Page / Section": f"Page {item['Page']} / {item['Section']}"
            })

    if not all_items:
        all_items.append({
            "Item No": "-",
            "Tender Work Item": "No BOQ Data Found",
            "Identified Work Type": "PDF format not supported or scanned PDF",
            "Quantity": "-",
            "Unit": "-",
            "Rate": "-",
            "Amount": "-",
            "Content Type": "REVIEW",
            "Related Tender Conditions": "-",
            "Technical Specification": "-",
            "Applicable Standard": "-",
            "Material Required": "-",
            "Approved Make": "-",
            "Vendor List": "-",
            "Execution Method": "-",
            "Tools & Equipment": "-",
            "Labour Requirement": "-",
            "Supervisor Requirement": "-",
            "Estimated Time": "-",
            "Risk / Review Point": "No BOQ rows detected. PDF may be scanned/image based or table layout unsupported.",
            "Confidence": "Low",
            "Source Page / Section": "-"
        })

    df = pd.DataFrame(all_items)

    file_name = f"vtenders_ai_output_{uuid.uuid4().hex}.xlsx"
    df.to_excel(file_name, index=False)

    format_excel(file_name)

    return FileResponse(
        path=file_name,
        filename=file_name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.get("/download/{file_name}")
def download_file(file_name: str):
    file_path = os.path.join(os.getcwd(), file_name)
    return FileResponse(path=file_path, filename=file_name)
