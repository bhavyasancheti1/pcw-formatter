# app.py
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import RedirectResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
import shutil
import os

from format_pcw import fill_template
# from format_pcw_json import fill_template_from_json  # keep if you use it elsewhere

app = FastAPI()

# --- CORS (allow your local Vite dev server + your Render domain) ---
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:5173",
        "http://localhost:5174",
        "https://pcw-formatter-2.onrender.com",
        # add your deployed frontend domain here when you have one
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Serve static/index.html as the home page (optional)
app.mount("/static", StaticFiles(directory="static"), name="static")

@app.get("/", include_in_schema=False)
def root():
    return RedirectResponse(url="/static/index.html")

@app.get("/healthz")
def healthz():
    return {"ok": True}

# ---- Generate PCW and return the file directly (no Google Drive) ----
@app.post("/generate-pcw/", summary="Generate a filled PCW Excel file and return as download")
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

        # Format the PCW file using your logic
        fill_template(gpt_output_path, pcw_template_path, output_path)

        if not os.path.exists(output_path):
            raise HTTPException(status_code=500, detail="Output file was not created")

        # Return the file directly so the browser downloads it
        return FileResponse(
            output_path,
            filename="Final_PCW_Filled.xlsx",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error generating PCW: {str(e)}")

