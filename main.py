from fastapi import FastAPI, UploadFile, File
from pptx import Presentation
import io

app = FastAPI()

@app.get("/")
def root():
    return {"status": "ok"}

def extract_shape_text(shape):
    texts = []

    if hasattr(shape, "text") and shape.text:
        texts.append(shape.text.strip())

    if hasattr(shape, "has_table") and shape.has_table:
        for row in shape.table.rows:
            cells = []
            for cell in row.cells:
                cell_text = cell.text.strip()
                cells.append(cell_text)
            texts.append(" | ".join(cells))

    if hasattr(shape, "shapes"):
        for sub_shape in shape.shapes:
            texts.extend(extract_shape_text(sub_shape))

    return texts

@app.post("/extract-ppt")
async def extract_ppt(file: UploadFile = File(...)):
    ppt_bytes = await file.read()
    prs = Presentation(io.BytesIO(ppt_bytes))

    result = []

    for i, slide in enumerate(prs.slides, start=1):
        texts = []

        for shape in slide.shapes:
            texts.extend(extract_shape_text(shape))

        cleaned_texts = []
        for t in texts:
            if t and t not in cleaned_texts:
                cleaned_texts.append(t)

        result.append({
            "slide": i,
            "text": "\n".join(cleaned_texts)
        })

    return result
