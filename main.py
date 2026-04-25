from fastapi import FastAPI, UploadFile, File
from fastapi.responses import HTMLResponse, FileResponse
from typing import List
import fitz
import pandas as pd
import uuid
import os

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

app = FastAPI()

# ---------------- KNOWLEDGE BASE ----------------
WORK_KNOWLEDGE_BASE = {
    "demolition": {
        "keywords": ["demolition", "dismantling", "dismentalling", "breaking", "removal", "dispose", "disposal"],
        "work": "Demolition / Dismantling Work",
        "materials": "No permanent material required; debris handling/consumables only",
        "method": "Site inspection, utility isolation, controlled breaking/dismantling, stacking of serviceable material, disposal of unserviceable material",
        "tools": "Breaker, hammer, chisel, cutter, wheelbarrow, safety barricade, PPE",
        "labour": "Demolition labour + supervisor",
        "forbidden": ["cement", "sand", "aggregate", "steel reinforcement"]
    },
    "concrete": {
        "keywords": ["concrete", "rcc", "pcc", "m20", "m25", "m30", "cement concrete"],
        "work": "Concrete Work",
        "materials": "Cement, sand, aggregate, water, admixture if specified, steel if RCC",
        "method": "Batching, mixing, placing, compaction, finishing and curing as per specification",
        "tools": "Concrete mixer, vibrator, cube mould, trowel, curing arrangement",
        "labour": "Mason + labour + supervisor",
        "forbidden": []
    },
    "brickwork": {
        "keywords": ["brick work", "brickwork", "masonry", "stone masonry", "block work"],
        "work": "Masonry Work",
        "materials": "Bricks/blocks/stone, cement mortar, sand, water",
        "method": "Line-level setting, mortar preparation, laying, joint filling and curing",
        "tools": "Trowel, line dori, level tube, plumb bob, scaffolding",
        "labour": "Mason + helper",
        "forbidden": []
    },
    "excavation": {
        "keywords": ["excavation", "earthwork", "digging", "trench", "pit excavation"],
        "work": "Excavation Work",
        "materials": "No permanent material; soil disposal/backfilling as applicable",
        "method": "Marking, excavation manually or by machine, dressing, dewatering if required, disposal/backfilling",
        "tools": "JCB, spade, pickaxe, dumper, measuring tape",
        "labour": "Excavator operator + labour + supervisor",
        "forbidden": ["cement", "sand", "aggregate"]
    },
    "cable_laying": {
        "keywords": ["cable", "cable laying", "xlpe", "lt cable", "ht cable", "cable pulling"],
        "work": "Cable Laying Work",
        "materials": "Cable, lugs, glands, saddles, tags, warning tape if specified",
        "method": "Route checking, drum handling, rollers placement, cable pulling, dressing, glanding, termination and testing",
        "tools": "Cable roller, drum jack, winch, crimping tool, megger, PPE",
        "labour": "Cable gang + electrician + supervisor",
        "forbidden": ["cement", "sand", "aggregate"]
    },
    "earthing": {
        "keywords": ["earthing", "earth pit", "earth electrode", "gi strip", "copper strip"],
        "work": "Earthing Work",
        "materials": "Earth electrode, GI/Cu strip, earth compound/charcoal/salt as specified, chamber cover",
        "method": "Pit excavation, electrode installation, strip connection, backfilling, watering and earth resistance testing",
        "tools": "Earth tester, spade, welding/drilling tools, multimeter",
        "labour": "Electrician + labour",
        "forbidden": []
    },
    "panel": {
        "keywords": ["panel", "switchgear", "mccb", "acb", "lt panel", "control panel"],
        "work": "Electrical Panel Work",
        "materials": "Panel, breakers, busbar, control wiring, lugs, glands, ferrules",
        "method": "Panel positioning, fixing, cable termination, control wiring, testing and commissioning",
        "tools": "Crimping tool, multimeter, torque wrench, insulation tester",
        "labour": "Electrician + technician + supervisor",
        "forbidden": ["cement", "sand", "aggregate"]
    },
    "scada": {
        "keywords": ["scada", "rtu", "plc", "hmi", "remote monitoring", "remote monitor", "automation"],
        "work": "SCADA / Remote Monitoring Work",
        "materials": "PLC/RTU, HMI, sensors, communication module, control cables, panel accessories",
        "method": "Panel wiring, device installation, communication setup, configuration, testing and commissioning",
        "tools": "Laptop, multimeter, crimping tool, communication tester",
        "labour": "Automation engineer + technician",
        "forbidden": ["cement", "sand", "aggregate"]
    },
    "painting": {
        "keywords": ["painting", "paint", "primer", "coating"],
        "work": "Painting / Coating Work",
        "materials": "Primer, paint/coating, thinner, putty if specified",
        "method": "Surface preparation, primer application, paint/coating application and finishing",
        "tools": "Brush, roller, spray gun, scraper, sandpaper",
        "labour": "Painter + helper",
        "forbidden": []
    }
}

# ---------------- PDF TEXT EXTRACTION ----------------
def extract_text_from_pdf(file):
    doc = fitz.open(stream=file, filetype="pdf")
    all_data = []

    for page in doc:
        blocks = page.get_text("blocks")
        for b in blocks:
            text = b[4].strip()
            if text:
                all_data.append(text)

    return all_data

# ---------------- CLEAN ----------------
def clean_lines(lines):
    return [line.replace("\n", " ").strip() for line in lines if line.strip()]

# ---------------- MERGE BROKEN LINES ----------------
def merge_lines(lines):
    merged = []
    buffer = ""

    for line in lines:
        line = line.strip()
        if line and line[0].isdigit():
            if buffer:
                merged.append(buffer.strip())
            buffer = line
        else:
            buffer += " " + line

    if buffer:
        merged.append(buffer.strip())

    return merged

# ---------------- BOQ EXTRACTION ----------------
def extract_boq(lines):
    boq = []

    for line in lines:
        parts = line.split()

        if len(parts) < 5 or not parts[0].isdigit():
            continue

        try:
            item_no = parts[0]

            numbers = []
            for p in parts:
                clean_p = p.replace(",", "").replace("₹", "")
                if clean_p.replace(".", "", 1).isdigit():
                    numbers.append(clean_p)

            if len(numbers) < 2:
                continue

            qty = numbers[-2]
            rate = numbers[-1]
            unit = parts[-1]

            description = " ".join(parts[1:-3])

            boq.append({
                "Item No": item_no,
                "Description": description,
                "Quantity": qty,
                "Rate": rate,
                "Unit": unit,
                "Source Line": line
            })
        except:
            continue

    return boq

# ---------------- SMART ANALYSIS ----------------
def analyze_work(description, unit):
    desc = description.lower()
    unit_l = str(unit).lower()

    matched_results = []

    for key, data in WORK_KNOWLEDGE_BASE.items():
        score = 0
        matched_keywords = []

        for kw in data["keywords"]:
            if kw in desc:
                score += 1
                matched_keywords.append(kw)

        if key == "concrete" and unit_l in ["cum", "cmt", "m3", "m³"]:
            score += 1

        if key == "cable_laying" and unit_l in ["m", "meter", "metre", "rmt"]:
            score += 1

        if score > 0:
            matched_results.append((score, key, data, matched_keywords))

    if not matched_results:
        return {
            "Work": "Unknown / Review Required",
            "Material": "-",
            "Method": "-",
            "Tools": "-",
            "Labour": "-",
            "Status": "⚠ Review Required",
            "Review Reason": "No matching work type found in knowledge base",
            "Confidence": "Low"
        }

    matched_results.sort(reverse=True, key=lambda x: x[0])
    best_score, best_key, best_data, matched_keywords = matched_results[0]

    confidence = "High" if best_score >= 2 else "Medium"

    status = "OK"
    review_reason = "-"

    material_lower = best_data["materials"].lower()
    for bad in best_data["forbidden"]:
        if bad in material_lower:
            status = "⚠ Review Required"
            review_reason = f"Forbidden material detected for {best_data['work']}: {bad}"

    if best_key == "demolition":
        if "m20" in desc or "m25" in desc or "rcc" in desc:
            status = "⚠ Review Required"
            review_reason = "Demolition item contains concrete/RCC terms; verify existing breaking or new concrete work"

    return {
        "Work": best_data["work"],
        "Material": best_data["materials"],
        "Method": best_data["method"],
        "Tools": best_data["tools"],
        "Labour": best_data["labour"],
        "Status": status,
        "Review Reason": review_reason,
        "Confidence": confidence
    }

# ---------------- RATE LOGIC ----------------
def get_rate(work_type, tender_rate):
    try:
        return float(tender_rate)
    except:
        return 0

# ---------------- EXCEL FORMATTING ----------------
def format_excel(file_name):
    wb = load_workbook(file_name)
    ws = wb.active

    ws.title = "Vtenders Output"

    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    header_fill = PatternFill(start_color="D9EAF7", end_color="D9EAF7", fill_type="solid")
    ok_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    review_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    medium_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    # Header format
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    # Data format
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = thin_border

    # Freeze header
    ws.freeze_panes = "A2"

    # Find Status and Confidence columns
    status_col = None
    confidence_col = None

    for cell in ws[1]:
        if cell.value == "Status":
            status_col = cell.column
        if cell.value == "Confidence":
            confidence_col = cell.column

    # Color status/confidence
    for row in range(2, ws.max_row + 1):
        if status_col:
            status_cell = ws.cell(row=row, column=status_col)
            value = str(status_cell.value)

            if value == "OK":
                status_cell.fill = ok_fill
                status_cell.font = Font(bold=True, color="006100")
            elif "Review" in value:
                status_cell.fill = review_fill
                status_cell.font = Font(bold=True, color="9C0006")

        if confidence_col:
            confidence_cell = ws.cell(row=row, column=confidence_col)
            value = str(confidence_cell.value)

            if value == "High":
                confidence_cell.fill = ok_fill
            elif value == "Medium":
                confidence_cell.fill = medium_fill
            elif value == "Low":
                confidence_cell.fill = review_fill

    # Column widths
    widths = {
        "A": 10,   # Item No
        "B": 28,   # Work
        "C": 55,   # Description
        "D": 45,   # Material
        "E": 12,   # Quantity
        "F": 12,   # Unit
        "G": 55,   # Method
        "H": 45,   # Tools
        "I": 35,   # Labour
        "J": 14,   # Rate
        "K": 16,   # Amount
        "L": 14,   # Confidence
        "M": 20,   # Status
        "N": 55,   # Review Reason
        "O": 70    # Source Line
    }

    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    # Row height
    for row in range(1, ws.max_row + 1):
        ws.row_dimensions[row].height = 45

    wb.save(file_name)

# ---------------- HOME UI ----------------
@app.get("/", response_class=HTMLResponse)
def home():
    return """
    <html>
    <body style="font-family: Arial; padding: 30px;">
        <h2>Vtenders - Tender File Analyzer</h2>
        <p>Upload multiple tender PDF files. System will extract BOQ and generate formatted Excel output.</p>

        <form action="/upload/" method="post" enctype="multipart/form-data">
            <input type="file" name="files" multiple required>
            <br><br>
            <button type="submit">Upload & Analyze</button>
        </form>
    </body>
    </html>
    """

# ---------------- MAIN PROCESS ----------------
@app.post("/upload/")
async def upload_files(files: List[UploadFile] = File(...)):
    all_items = []

    for file in files:
        content = await file.read()

        lines = extract_text_from_pdf(content)
        lines = clean_lines(lines)
        lines = merge_lines(lines)
        boq = extract_boq(lines)

        for item in boq:
            analysis = analyze_work(item["Description"], item["Unit"])

            try:
                qty = float(str(item["Quantity"]).replace(",", ""))
            except:
                qty = 0

            rate = get_rate(analysis["Work"], item["Rate"])
            amount = qty * rate

            all_items.append({
                "Item No": item["Item No"],
                "Work": analysis["Work"],
                "Description": item["Description"],
                "Material": analysis["Material"],
                "Quantity": qty,
                "Unit": item["Unit"],
                "Method": analysis["Method"],
                "Tools": analysis["Tools"],
                "Labour": analysis["Labour"],
                "Rate": rate,
                "Amount": amount,
                "Confidence": analysis["Confidence"],
                "Status": analysis["Status"],
                "Review Reason": analysis["Review Reason"],
                "Source Line": item["Source Line"]
            })

    if not all_items:
        all_items.append({
            "Item No": "-",
            "Work": "No Data Found",
            "Description": "PDF format not supported or BOQ table not detected",
            "Material": "-",
            "Quantity": "-",
            "Unit": "-",
            "Method": "-",
            "Tools": "-",
            "Labour": "-",
            "Rate": "-",
            "Amount": "-",
            "Confidence": "Low",
            "Status": "⚠ Review Required",
            "Review Reason": "No BOQ rows detected",
            "Source Line": "-"
        })

    df = pd.DataFrame(all_items)

    file_name = f"vtenders_output_{uuid.uuid4().hex}.xlsx"
    df.to_excel(file_name, index=False)

    format_excel(file_name)

    return FileResponse(
        path=file_name,
        filename=file_name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ---------------- DOWNLOAD ----------------
@app.get("/download/{file_name}")
def download_file(file_name: str):
    file_path = os.path.join(os.getcwd(), file_name)
    return FileResponse(path=file_path, filename=file_name)