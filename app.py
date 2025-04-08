# app.py

from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
import shutil
from format_pcw import fill_template

app = FastAPI()

@app.post("/generate-pcw/")
async def generate_pcw(gpt_output: UploadFile = File(...), pcw_template: UploadFile = File(...)):
    gpt_output_path = f"/tmp/{gpt_output.filename}"
    pcw_template_path = f"/tmp/{pcw_template.filename}"
    output_path = "/tmp/Final_PCW_Filled_With_All_Labors.xlsx"

    with open(gpt_output_path, "wb") as f:
        shutil.copyfileobj(gpt_output.file, f)
    with open(pcw_template_path, "wb") as f:
        shutil.copyfileobj(pcw_template.file, f)

    fill_template(gpt_output_path, pcw_template_path, output_path)
    return FileResponse(output_path, filename="Final_PCW.xlsx")
