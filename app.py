# app.py


from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
import shutil
import os
from format_pcw import fill_template

app = FastAPI()

@app.post("/generate-pcw/", summary="Generate a filled PCW Excel file from GPT output and template")
async def generate_pcw(
    gpt_output: UploadFile = File(...),
    pcw_template: UploadFile = File(...)
):
    # Validate file types
    if not gpt_output.filename.endswith(('.xlsx', '.xls')) or not pcw_template.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=422, detail="Both files must be Excel (.xlsx or .xls)")

    # Save files
    gpt_output_path = f"/tmp/{gpt_output.filename}"
    pcw_template_path = f"/tmp/{pcw_template.filename}"
    output_path = "/tmp/Final_PCW_Auto.xlsx"

    try:
        with open(gpt_output_path, 'wb') as f:
            shutil.copyfileobj(gpt_output.file, f)
        with open(pcw_template_path, 'wb') as f:
            shutil.copyfileobj(pcw_template.file, f)

        # Call your core formatter logic
        fill_template(gpt_output_path, pcw_template_path, output_path)

        # Return the filled Excel
        return FileResponse(
            output_path,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename="Final_PCW_Filled.xlsx"
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error generating PCW: {str(e)}")
