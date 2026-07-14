"""
main.py — FastAPI бэкенд для проверки деловых писем (.docx)

Запуск (Windows):
    python -m uvicorn main:app --reload --port 8000
"""

import shutil
import tempfile
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from pydantic import BaseModel

from extract_docx_header import check_plain_text, extract_header_and_greeting

app = FastAPI(title="LetterCheck Smart API", version="2.0.0")


class TextCheckRequest(BaseModel):
    text: str

# Настройка CORS для фронтенда
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173", "http://127.0.0.1:5173"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/check")
async def check_document(file: UploadFile = File(...)):
    """Проверка загруженного .docx файла"""
    if not file.filename or not file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Принимаются только файлы .docx")

    tmp_dir = Path(tempfile.mkdtemp())
    tmp_path = tmp_dir / file.filename

    try:
        with open(tmp_path, "wb") as f:
            shutil.copyfileobj(file.file, f)

        result = extract_header_and_greeting(str(tmp_path))
        return JSONResponse(content=result)

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ошибка при обработке файла: {e}")

    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


@app.post("/check-text")
async def check_text(payload: TextCheckRequest):
    """Проверка обычного вставленного текста."""
    if not payload.text or not payload.text.strip():
        raise HTTPException(status_code=400, detail="Текст для проверки пустой")
    return JSONResponse(content=check_plain_text(payload.text))


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000, reload=True)
    
