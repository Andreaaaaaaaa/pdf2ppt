from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import StreamingResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from converter import PDFToPPTConverter
import io

app = FastAPI()

@app.post("/convert")
async def convert_pdf(
    file: UploadFile = File(...),
    mode: str = Form(...),
    dpi: int = Form(200)
):
    try:
        # Read file content
        content = await file.read()
        pdf_stream = io.BytesIO(content)
        
        converter = PDFToPPTConverter(pdf_stream)
        
        if mode == "separated":
            ppt_stream = converter.convert_separated()
        else:
            ppt_stream = converter.convert_to_images(dpi=dpi)
            
        return StreamingResponse(
            ppt_stream,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={"Content-Disposition": "attachment; filename=converted.pptx"}
        )
    except Exception as e:
        import traceback
        traceback.print_exc()
        return JSONResponse(
            status_code=500,
            content={"detail": f"Text extraction error: {str(e)}"}
        )

@app.post("/extract_text")
async def extract_text(file: UploadFile = File(...)):
    try:
        content = await file.read()
        pdf_stream = io.BytesIO(content)
        
        converter = PDFToPPTConverter(pdf_stream)
        text_stream = converter.extract_text_content()
        
        return StreamingResponse(
            text_stream,
            media_type="text/plain",
            headers={"Content-Disposition": "attachment; filename=extracted_text.txt"}
        )
    except Exception as e:
        import traceback
        traceback.print_exc()
        return JSONResponse(
            status_code=500,
            content={"detail": f"Text extraction error: {str(e)}"}
        )

# Mount static files (after API routes to avoid interception)
app.mount("/", StaticFiles(directory="web", html=True), name="static")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
