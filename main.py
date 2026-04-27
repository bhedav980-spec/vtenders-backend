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

# ---------------- PDF TEXT ----------------
def extract_text_from_pdf(file):
    doc = fitz.open(stream=file, filetype="pdf")
    lines = []

    for page in doc:
        for line in page.get_text("text").splitlines():
            line = re.sub(r"\s+", " ", line.strip())
            if line:
                lines.append(line)

    return lines

# ---------------- BOQ EXTRACTION ----------------
def extract_boq(lines):
    boq = []
    pattern = re.compile(r"^\s*(\d+)\s+(.*?)\s+(\d+(\.\d+)?)\s+(\w+)\s+(\d+(\.\d+)?)")

    for line in lines:
        m = pattern.search(line)
        if m:
            boq.append({
                "Item No": m.group(1),
                "Description": m.group(2),
                "Quantity": m.group(3),
                "Unit": m.group(5),
                "Rate": m.group(6),
                "Source": line
            })

    return boq

# ---------------- AI ANALYSIS ----------------
def ai_analyze(description):
    try:
        prompt = f"""
You are a senior tender engineer.

Understand the work and return JSON:

{{
"Work":"",
"Material":"",
"Method":"",
"Tools":"",
"Labour":"",
"Confidence":""
}}

TEXT:
{description}
"""

        res = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role":"system","content":"Return only JSON"},
                {"role":"user","content":prompt}
            ],
            temperature=0
        )

        data = json.loads(res.choices[0].message.content)
        return data

    except:
        return {
            "Work":"General Work",
            "Material":"-",
            "Method":"-",
            "Tools":"-",
            "Labour":"-",
            "Confidence":"Low"
        }

# ---------------- EXCEL FORMAT (OLD STYLE) ----------------
def format_excel(file):
    wb = load_workbook(file)
    ws = wb.active

    thin = Border(left=Side(style="thin"), right=Side(style="thin"),
                  top=Side(style="thin"), bottom=Side(style="thin"))

    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(wrap_text=True)
        cell.border = thin

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True)
            cell.border = thin

    wb.save(file)

# ---------------- ROUTES ----------------
@app.get("/", response_class=HTMLResponse)
def home():
    return "<h2>Vtenders Running</h2>"

@app.post("/upload/")
async def upload(files: List[UploadFile] = File(...)):
    output = []

    for f in files:
        content = await f.read()
        lines = extract_text_from_pdf(content)
        boq = extract_boq(lines)

        for item in boq:
            ai = ai_analyze(item["Description"])

            output.append({
                "Item No": item["Item No"],
                "Work": ai["Work"],
                "Description": item["Description"],
                "Material": ai["Material"],
                "Quantity": item["Quantity"],
                "Unit": item["Unit"],
                "Method": ai["Method"],
                "Tools": ai["Tools"],
                "Labour": ai["Labour"],
                "Rate": item["Rate"],
                "Confidence": ai["Confidence"]
            })

    df = pd.DataFrame(output)

    file_name = f"vtenders_output_{uuid.uuid4().hex}.xlsx"
    df.to_excel(file_name, index=False)

    format_excel(file_name)

    return FileResponse(file_name)
