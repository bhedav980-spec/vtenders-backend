from fastapi import FastAPI, UploadFile, File
from fastapi.responses import HTMLResponse, FileResponse
from typing import List
import fitz
import pandas as pd
import uuid
import os
import re

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

app = FastAPI()

WORK_KNOWLEDGE_BASE = {
    "demolition": {
        "keywords": ["demolition", "dismantling", "dismantle", "dismentalling", "breaking", "removal", "remove", "dispose", "disposal", "scrap"],
        "work": "Demolition / Dismantling Work",
        "materials": "No permanent material required; debris handling and consumables only",
        "method": "Site inspection, utility isolation, controlled dismantling/breaking, stacking of serviceable material and disposal of unserviceable material",
        "tools": "Breaker, hammer, chisel, cutter, wheelbarrow, safety barricade, PPE",
        "labour": "Demolition labour + supervisor",
    },
    "concrete": {
        "keywords": ["concrete", "rcc", "pcc", "m10", "m15", "m20", "m25", "m30", "cement concrete", "reinforced cement concrete"],
        "work": "Concrete Work",
        "materials": "Cement, sand, aggregate, water, admixture if specified, steel if RCC",
        "method": "Batching, mixing, placing, compaction, finishing and curing as per specification",
        "tools": "Concrete mixer, vibrator, cube mould, trowel, curing arrangement",
        "labour": "Mason + labour + supervisor",
    },
    "brickwork": {
        "keywords": ["brick work", "brickwork", "masonry", "stone masonry", "block work", "aac block", "fly ash brick"],
        "work": "Masonry Work",
        "materials": "Bricks/blocks/stone, cement mortar, sand and water",
        "method": "Line-level setting, mortar preparation, laying, joint filling and curing",
        "tools": "Trowel, line dori, level tube, plumb bob, scaffolding",
        "labour": "Mason + helper",
    },
    "excavation": {
        "keywords": ["excavation", "earthwork", "earth work", "digging", "trench", "pit excavation", "soil", "backfilling", "back filling"],
        "work": "Excavation / Earthwork",
        "materials": "No permanent material; soil disposal/backfilling as applicable",
        "method": "Marking, excavation manually or by machine, dressing, dewatering if required, disposal/backfilling",
        "tools": "JCB, spade, pickaxe, dumper, measuring tape",
        "labour": "Excavator operator + labour + supervisor",
    },
    "cable_laying": {
        "keywords": ["cable", "cable laying", "xlpe", "lt cable", "ht cable", "cable pulling", "armoured cable", "unarmoured cable"],
        "work": "Cable Laying Work",
        "materials": "Cable, lugs, glands, saddles, tags, warning tape if specified",
        "method": "Route checking, drum handling, rollers placement, cable pulling, dressing, glanding, termination and testing",
        "tools": "Cable roller, drum jack, winch, crimping tool, megger, PPE",
        "labour": "Cable gang + electrician + supervisor",
    },
    "earthing": {
        "keywords": ["earthing", "earth pit", "earth electrode", "gi strip", "copper strip", "earth mat", "chemical earthing"],
        "work": "Earthing Work",
        "materials": "Earth electrode, GI/Cu strip, earth compound/charcoal/salt as specified, chamber cover",
        "method": "Pit excavation, electrode installation, strip connection, backfilling, watering and earth resistance testing",
        "tools": "Earth tester, spade, welding/drilling tools, multimeter",
        "labour": "Electrician + labour",
    },
    "panel": {
        "keywords": ["panel", "switchgear", "mccb", "acb", "lt panel", "control panel", "db", "distribution board", "feeder pillar"],
        "work": "Electrical Panel Work",
        "materials": "Panel, breakers, busbar, control wiring, lugs, glands, ferrules",
        "method": "Panel positioning, fixing, cable termination, control wiring, testing and commissioning",
        "tools": "Crimping tool, multimeter, torque wrench, insulation tester",
        "labour": "Electrician + technician + supervisor",
    },
    "scada": {
        "keywords": ["scada", "rtu", "plc", "hmi", "remote monitoring", "remote monitor", "automation", "data logger", "communication"],
        "work": "SCADA / Remote Monitoring Work",
        "materials": "PLC/RTU, HMI, sensors, communication module, control cables, panel accessories",
        "method": "Panel wiring, device installation, communication setup, configuration, testing and commissioning",
        "tools": "Laptop, multimeter, crimping tool, communication tester",
        "labour": "Automation engineer + technician",
    },
    "painting": {
        "keywords": ["painting", "paint", "primer", "coating", "enamel", "epoxy", "white wash", "distemper"],
        "work": "Painting / Coating Work",
        "materials": "Primer, paint/coating, thinner, putty if specified",
        "method": "Surface preparation, primer application, paint/coating application and finishing",
        "tools": "Brush, roller, spray gun, scraper, sandpaper",
        "labour": "Painter + helper",
    },
    "plaster": {
        "keywords": ["plaster", "plastering", "cement plaster", "12mm plaster", "15mm plaster", "20mm plaster"],
        "work": "Plaster Work",
        "materials": "Cement, sand, water and curing arrangement",
        "method": "Surface preparation, mortar mixing, plaster application, levelling, finishing and curing",
        "tools": "Trowel, level, wooden float, scaffolding",
        "labour": "Mason + helper",
    },
    "flooring": {
        "keywords": ["flooring", "tile", "tiles", "vitrified", "ceramic", "granite", "marble", "paver block"],
        "work": "Flooring / Tiling Work",
        "materials": "Tiles/stone/paver, adhesive or mortar, grout, cement, sand and water",
        "method": "Surface preparation, line-level marking, laying, joint filling, cleaning and curing",
        "tools": "Tile cutter, trowel, level, spacer, rubber mallet",
        "labour": "Tile mason + helper",
    },
    "fabrication": {
        "keywords": ["fabrication", "structural steel", "ms steel", "steel structure", "welding", "ms angle", "ms channel", "truss"],
        "work": "Fabrication Work",
        "materials": "MS steel sections, welding electrodes, bolts, primer/paint as specified",
        "method": "Cutting, fitting, welding/bolting, alignment, finishing and painting",
        "tools": "Welding machine, grinder, gas cutter, drill machine, measuring tools",
        "labour": "Fabricator + welder + helper",
    },
    "solar": {
        "keywords": ["solar", "module", "pv module", "inverter", "string", "solar plant", "mms", "dc cable", "ac cable"],
        "work": "Solar Plant Work",
        "materials": "Solar modules, MMS, inverter, cables, connectors, earthing and accessories as specified",
        "method": "Installation, alignment, cabling, termination, testing and commissioning",
        "tools": "Torque wrench, multimeter, megger, MC4 crimper, PPE",
        "labour": "Solar technician + electrician + supervisor",
    }
}


def extract_text_from_pdf(file):
    doc = fitz.open(stream=file, filetype="pdf")
    all_lines = []

    for page in doc:
        text = page.get_text("text")
        for line in text.splitlines():
            line = line.strip()
            if line:
                all_lines.append(line)

    return all_lines


def clean_lines(lines):
    cleaned = []
    for line in lines:
        line = re.sub(r"\s+", " ", line.replace("\n", " ")).strip()
        if line:
            cleaned.append(line)
    return cleaned


def merge_lines(lines):
    merged = []
    buffer = ""

    item_start = re.compile(r"^\s*(\d+|[0-9]+\.[0-9]+)\s+")

    for line in lines:
        if item_start.match(line):
            if buffer:
                merged.append(buffer.strip())
            buffer = line
        else:
            if buffer:
                buffer += " " + line
            else:
                buffer = line

    if buffer:
        merged.append(buffer.strip())

    return merged


def is_number(value):
    value = str(value).replace(",", "").replace("₹", "").strip()
    return bool(re.fullmatch(r"\d+(\.\d+)?", value))


def normalize_unit(unit):
    u = str(unit).strip().lower()
    unit_map = {
        "cum": "Cum", "cu.m": "Cum", "cu.m.": "Cum", "m3": "Cum", "m³": "Cum", "cmt": "Cum",
        "sqm": "Sqm", "sq.m": "Sqm", "sq.m.": "Sqm", "m2": "Sqm", "m²": "Sqm", "smt": "Sqm",
        "rm": "Rmt", "rmt": "Rmt", "m": "Mtr", "meter": "Mtr", "metre": "Mtr", "mtr": "Mtr",
        "nos": "Nos", "no": "Nos", "each": "Nos", "set": "Set", "job": "Job",
        "kg": "Kg", "mt": "MT", "ton": "MT", "ltr": "Ltr", "litre": "Ltr"
    }
    return unit_map.get(u, unit)


def extract_boq(lines):
    boq = []
    known_units = r"(cum|cu\.m\.?|m3|m³|cmt|sqm|sq\.m\.?|m2|m²|smt|rmt|rm|mtr|meter|metre|m|nos|no|each|set|job|kg|mt|ton|ltr|litre)"

    pattern = re.compile(
        rf"^\s*(?P<item>\d+(\.\d+)*)\s+"
        rf"(?P<desc>.*?)\s+"
        rf"(?P<qty>\d+(\.\d+)?)\s*"
        rf"(?P<unit>{known_units})\s+"
        rf"(?P<rate>\d+(\.\d+)?)",
        re.IGNORECASE
    )

    for line in lines:
        line_clean = line.replace(",", "")
        match = pattern.search(line_clean)

        if match:
            boq.append({
                "Item No": match.group("item"),
                "Description": match.group("desc").strip(),
                "Quantity": match.group("qty"),
                "Rate": match.group("rate"),
                "Unit": normalize_unit(match.group("unit")),
                "Source Line": line
            })
            continue

        parts = line_clean.split()
        if len(parts) < 6:
            continue

        if not re.match(r"^\d+(\.\d+)*$", parts[0]):
            continue

        unit_index = None
        for i, p in enumerate(parts):
            if re.fullmatch(known_units, p.lower()):
                unit_index = i
                break

        if unit_index and unit_index >= 2:
            qty_index = unit_index - 1
            rate_index = unit_index + 1

            if rate_index < len(parts) and is_number(parts[qty_index]) and is_number(parts[rate_index]):
                boq.append({
                    "Item No": parts[0],
                    "Description": " ".join(parts[1:qty_index]).strip(),
                    "Quantity": parts[qty_index],
                    "Rate": parts[rate_index],
                    "Unit": normalize_unit(parts[unit_index]),
                    "Source Line": line
                })

    return boq


def analyze_work(description, unit):
    desc = description.lower()
    unit_l = str(unit).lower()

    best_score = 0
    best_data = None
    matched_keywords = []

    for _, data in WORK_KNOWLEDGE_BASE.items():
        score = 0
        found = []

        for kw in data["keywords"]:
            if kw in desc:
                score += 3
                found.append(kw)

        if unit_l in ["cum"] and any(x in desc for x in ["concrete", "excavation", "earthwork", "pcc", "rcc"]):
            score += 2

        if unit_l in ["sqm"] and any(x in desc for x in ["painting", "plaster", "flooring", "tile"]):
            score += 2

        if unit_l in ["mtr", "rmt"] and any(x in desc for x in ["cable", "pipe", "strip"]):
            score += 2

        if score > best_score:
            best_score = score
            best_data = data
            matched_keywords = found

    if best_data is None:
        return {
            "Work": "General Tender Item",
            "Material": "As per item description / tender specification",
            "Method": "Execute work as per tender specification, drawing and site instruction",
            "Tools": "Standard tools, tackles and safety equipment as required",
            "Labour": "Skilled/semi-skilled labour + supervisor as required",
            "Status": "OK",
            "Review Reason": "-",
            "Confidence": "Low"
        }

    confidence = "High" if best_score >= 5 else "Medium"

    return {
        "Work": best_data["work"],
        "Material": best_data["materials"],
        "Method": best_data["method"],
        "Tools": best_data["tools"],
        "Labour": best_data["labour"],
        "Status": "OK",
        "Review Reason": "-",
        "Confidence": confidence
    }


def get_rate(tender_rate):
    try:
        return float(str(tender_rate).replace(",", "").replace("₹", ""))
    except:
        return 0


def format_excel(file_name):
    wb = load_workbook(file_name)
    ws = wb.active
    ws.title = "Vtenders Output"

    thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

    header_fill = PatternFill(start_color="0B3D91", end_color="0B3D91", fill_type="solid")
    ok_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    medium_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    low_fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")

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
    status_col = None

    for cell in ws[1]:
        if cell.value == "Confidence":
            confidence_col = cell.column
        if cell.value == "Status":
            status_col = cell.column

    for row in range(2, ws.max_row + 1):
        if confidence_col:
            c = ws.cell(row=row, column=confidence_col)
            if c.value == "High":
                c.fill = ok_fill
            elif c.value == "Medium":
                c.fill = medium_fill
            else:
                c.fill = low_fill

        if status_col:
            s = ws.cell(row=row, column=status_col)
            s.fill = ok_fill
            s.font = Font(bold=True, color="006100")

    widths = {
        "A": 10, "B": 30, "C": 65, "D": 48, "E": 12, "F": 12,
        "G": 60, "H": 45, "I": 35, "J": 14, "K": 16,
        "L": 14, "M": 16, "N": 45, "O": 75
    }

    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    for row in range(1, ws.max_row + 1):
        ws.row_dimensions[row].height = 42

    wb.save(file_name)


@app.get("/", response_class=HTMLResponse)
def home():
    return """
    <html>
    <body style="font-family: Arial; padding: 30px;">
        <h2>Vtenders - Tender File Analyzer</h2>
        <p>Upload tender PDF files. System will extract BOQ and generate Excel output.</p>
        <form action="/upload/" method="post" enctype="multipart/form-data">
            <input type="file" name="files" multiple required>
            <br><br>
            <button type="submit">Upload & Analyze</button>
        </form>
    </body>
    </html>
    """


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

            rate = get_rate(item["Rate"])
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
            "Work": "No BOQ Data Found",
            "Description": "BOQ table not detected. PDF may be scanned/image based or layout is unsupported.",
            "Material": "-",
            "Quantity": "-",
            "Unit": "-",
            "Method": "-",
            "Tools": "-",
            "Labour": "-",
            "Rate": "-",
            "Amount": "-",
            "Confidence": "Low",
            "Status": "Review",
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


@app.get("/download/{file_name}")
def download_file(file_name: str):
    file_path = os.path.join(os.getcwd(), file_name)
    return FileResponse(path=file_path, filename=file_name)
