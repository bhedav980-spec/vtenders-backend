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

from openai import OpenAI

app = FastAPI()

client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

# ---------------- PDF EXTRACTION ----------------
def extract_text_from_pdf(file_bytes):
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    lines = []

    for page_no, page in enumerate(doc, start=1):
        text = page.get_text("text")
        for line in text.splitlines():
            line = re.sub(r"\s+", " ", line).strip()
            if line:
                lines.append({"page": page_no, "text": line})

    return lines

# ---------------- SECTION DETECTION ----------------
def detect_section(line):
    t = line.lower()

    if "boq" in t or "bill of quantity" in t:
        return "BOQ"
    if "scope" in t:
        return "SCOPE_OF_WORK"
    if "specification" in t:
        return "TECHNICAL_SPECIFICATION"
    if "approved make" in t:
        return "APPROVED_MAKE"
    if "vendor" in t:
        return "VENDOR_LIST"
    if "condition" in t:
        return "CONDITION"
    if "safety" in t:
        return "SAFETY"
    if "drawing" in t:
        return "DRAWING"

    return "OTHER"

def build_context(lines):
    current = "OTHER"
    out = []

    for row in lines:
        sec = detect_section(row["text"])
        if sec != "OTHER":
            current = sec

        out.append({
            "page": row["page"],
            "section": current,
            "text": row["text"]
        })

    return out

# ---------------- CONTEXT MATCHING ----------------
def get_relevant_context(description, context_rows, max_lines=30):
    words = set(re.findall(r"[a-zA-Z]{4,}", description.lower()))
    scored = []

    for row in context_rows:
        score = 0
        text = row["text"].lower()

        for w in words:
            if w in text:
                score += 2

        if row["section"] != "OTHER":
            score += 1

        if score > 0:
            scored.append((score, row))

    scored.sort(reverse=True, key=lambda x: x[0])
    top = scored[:max_lines]

    return "\n".join([f"[Page {r['page']} | {r['section']}] {r['text']}" for _, r in top])

# ---------------- BOQ EXTRACTION ----------------
def extract_boq(context_rows):
    boq = []
    pattern = re.compile(r"^\s*(\d+)\s+(.*?)\s+(\d+(\.\d+)?)\s+(\w+)\s+(\d+(\.\d+)?)")

    for row in context_rows:
        line = row["text"]
        m = pattern.search(line)

        if m:
            boq.append({
                "Item No": m.group(1),
                "Description": m.group(2),
                "Quantity": m.group(3),
                "Unit": m.group(5),
                "Rate": m.group(6),
                "Page": row["page"],
                "Section": row["section"]
            })

    return boq

# ---------------- AI ANALYSIS ----------------
def safe_json(text):
    try:
        return json.loads(text)
    except:
        return None

def ai_analyze(item, context):
    prompt = f"""
You are a senior EPC tender analysis engineer.

TASK:
Understand BOQ item and tender deeply.

Extract:
- Work type
- Conditions
- Technical specs
- Standards
- Approved make
- Vendor
- Method
- Tools
- Labour
- Time

Return ONLY JSON:

{{
"content_type":"",
"identified_work_type":"",
"related_conditions":"",
"technical_specification":"",
"applicable_standard":"",
"material_required":"",
"approved_make":"",
"vendor_list":"",
"execution_method":"",
"tools_equipment":"",
"labour_requirement":"",
"supervisor_requirement":"",
"estimated_time":"",
"risk_review_point":"",
"confidence":""
}}

ITEM:
{item["Description"]}

CONTEXT:
{context}
"""

    try:
        res = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role":"system","content":"Return only JSON"},
                {"role":"user","content":prompt}
            ],
            temperature=0
        )

        data = safe_json(res.choices[0].message.content)

        if data:
            return data
        else:
            return {"identified_work_type":"Unknown","confidence":"Low"}

    except Exception as e:
        return {"identified_work_type":"Error","risk_review_point":str(e),"confidence":"Low"}

# ---------------- EXCEL ----------------
def format_excel(file):
    wb = load_workbook(file)
    ws = wb.active

    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(wrap_text=True)

    wb.save(file)

# ---------------- ROUTES ----------------
@app.get("/", response_class=HTMLResponse)
def home():
    return "<h2>Vtenders AI Running</h2>"

@app.get("/ai-test")
def test():
    return {"ai_enabled": True}

@app.post("/upload/")
async def upload(files: List[UploadFile] = File(...)):
    output = []

    for f in files:
        content = await f.read()

        lines = extract_text_from_pdf(content)
        context = build_context(lines)
        boq = extract_boq(context)

        for item in boq:
            ctx = get_relevant_context(item["Description"], context)
            ai = ai_analyze(item, ctx)

            output.append({
                "Item No": item["Item No"],
                "Work": ai.get("identified_work_type"),
                "Description": item["Description"],
                "Material": ai.get("material_required"),
                "Method": ai.get("execution_method"),
                "Tools": ai.get("tools_equipment"),
                "Labour": ai.get("labour_requirement"),
                "Time": ai.get("estimated_time"),
                "Confidence": ai.get("confidence")
            })

    df = pd.DataFrame(output)
    name = f"vtenders_ai_{uuid.uuid4().hex}.xlsx"
    df.to_excel(name, index=False)
    format_excel(name)

    return FileResponse(name)
