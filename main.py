import os
import shutil
import uuid
import tempfile
from typing import List
from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pdf_to_excel import build_excel

app = FastAPI(title="PDF to Excel Converter")

# Mount static files
app.mount("/static", StaticFiles(directory="static"), name="static")

@app.get("/")
async def read_index():
    return FileResponse("static/index.html")

def cleanup_files(temp_dir: str):
    """Removes the temporary directory after the file has been sent."""
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)

@app.post("/convert")
async def convert_pdfs(background_tasks: BackgroundTasks, files: List[UploadFile] = File(...)):
    if not files:
        raise HTTPException(status_code=400, detail="No files uploaded")

    # Create a unique temporary directory for this request
    temp_dir = tempfile.mkdtemp()
    pdf_paths = []
    
    try:
        for file in files:
            if not file.filename.lower().endswith(".pdf"):
                continue
            
            # Save uploaded file to temp directory
            file_path = os.path.join(temp_dir, file.filename)
            with open(file_path, "wb") as buffer:
                shutil.copyfileobj(file.file, buffer)
            pdf_paths.append(file_path)

        if not pdf_paths:
            raise HTTPException(status_code=400, detail="No valid PDF files uploaded")

        output_filename = f"extracted_data_{uuid.uuid4().hex[:8]}.xlsx"
        output_path = os.path.join(temp_dir, output_filename)

        # Run the conversion logic
        # Note: build_excel is imported from pdf_to_excel.py
        build_excel(pdf_paths, output_path=output_path)

        if not os.path.exists(output_path):
            raise HTTPException(status_code=500, detail="Conversion failed to generate file")

        # Schedule cleanup of the temp directory
        background_tasks.add_task(cleanup_files, temp_dir)

        return FileResponse(
            path=output_path,
            filename="extracted_data.xlsx",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        cleanup_files(temp_dir)
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
