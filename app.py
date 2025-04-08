# app.py

from fastapi import FastAPI, UploadFile, File, HTTPException, Request
from fastapi.responses import FileResponse
import shutil
import os
import tempfile
import requests

from format_pcw import fill_template
from format_pcw_json import fill_template_from_json

app = FastAPI()

# === Existing Endpoint: File Upload ===
@app.post("/generate-pcw/", summary="Generate a filled PCW Excel file from uploaded GPT output and template")
async def generate_pcw(
    gpt_output: UploadFile = File(...),
    pcw_template: UploadFile = File(...)
):
    print("✅ /generate-pcw/ endpoint was hit")
    print(f"📂 Received files: {gpt_output.filename}, {pcw_template.filename}")

    if not gpt_output.filename.endswith(('.xlsx', '.xls')) or not pcw_template.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=422, detail="Both files must be Excel (.xlsx or .xls)")

    gpt_output_path = f"/tmp/{gpt_output.filename}"
    pcw_template_path = f"/tmp/{pcw_template.filename}"
    output_path = "/tmp/Final_PCW_Filled.xlsx"

    try:
        with open(gpt_output_path, 'wb') as f:
            shutil.copyfileobj(gpt_output.file, f)
        with open(pcw_template_path, 'wb') as f:
            shutil.copyfileobj(pcw_template.file, f)

        fill_template(gpt_output_path, pcw_template_path, output_path)

        return FileResponse(output_path, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename="Final_PCW_Filled.xlsx")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error generating PCW: {str(e)}")

# === New Endpoint: JSON Input for GPT Integration ===
@app.post("/format-pcw-json/")
async def format_pcw_json(request: Request):
    try:
        body = await request.json()
        gpt_data = body.get("gpt_data")
        pcw_template_url = body.get("pcw_template_url")

        if not gpt_data or not pcw_template_url:
            raise HTTPException(status_code=400, detail="Missing gpt_data or pcw_template_url")

        template_response = requests.get(pcw_template_url)
        if template_response.status_code != 200:
            raise HTTPException(status_code=400, detail="Failed to fetch PCW template file")

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_template:
            temp_template.write(template_response.content)
            temp_template_path = temp_template.name

        output_path = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name

        fill_template_from_json(gpt_data, temp_template_path, output_path)

        return FileResponse(output_path, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename="Final_PCW_Generated.xlsx")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
