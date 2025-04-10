# app.py

from fastapi import FastAPI, UploadFile, File, HTTPException, Request
from fastapi.responses import JSONResponse
import shutil
import os
import tempfile
import requests

from format_pcw import fill_template
from format_pcw_json import fill_template_from_json
from gdrive_uploader import upload_file_to_gdrive  # 👈 NEW

app = FastAPI()

# === Updated Endpoint: File Upload + Google Drive Upload ===
@app.post("/generate-pcw/", summary="Generate a filled PCW Excel file and upload to Drive")
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
        # Save the uploaded files
        with open(gpt_output_path, 'wb') as f:
            shutil.copyfileobj(gpt_output.file, f)
        with open(pcw_template_path, 'wb') as f:
            shutil.copyfileobj(pcw_template.file, f)

        # Call your formatting function
        fill_template(gpt_output_path, pcw_template_path, output_path)

        # Upload to Google Drive
        view_url, download_url = upload_file_to_gdrive(output_path, "Final_PCW_Filled.xlsx")

        # Return the Drive links instead of a file response
        return {
            "message": "✅ PCW file uploaded to Google Drive.",
            "view_link": view_url,
            "download_link": download_url
        }

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error generating PCW: {str(e)}")
