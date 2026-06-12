from fastapi import FastAPI, UploadFile, File
from pptx import Presentation
import tempfile

app = FastAPI()

@app.get("/")
def root():
    return {"status":"ok"}

@app.post("/extract-ppt")
async def extract_ppt(file: UploadFile = File(...)):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
        tmp.write(await file.read())
        tmp_path = tmp.name

    prs = Presentation(tmp_path)

    result = []

    for i, slide in enumerate(prs.slides, start=1):
        texts = []

        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text:
                texts.append(shape.text)

        result.append({
            "slide": i,
            "text": "\n".join(texts)
        })

    return result
