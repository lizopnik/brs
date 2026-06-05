"""
Сервер для Яндекс Форм: приём ответа формы и генерация Word-документа БРС.

На каждый POST от интеграции Яндекс Форм сервер:
  1. сохраняет «сырой» запрос (тело + заголовки) в captures/ — для отладки;
  2. парсит ответ и заполняет шаблон data/*.docx, сохраняя готовый
     БРС_Фамилия_Имя_Отчество.docx в output/;
  3. отвечает кодом 200 (этого Яндексу достаточно, иначе он повторяет запрос).

Запуск:
    uvicorn main:app --host 0.0.0.0 --port 8000 --reload

Затем в отдельном окне поднять туннель Cloudflare (см. README.md):
    cloudflared tunnel --url http://localhost:8000
"""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse

from brs_generator import generate_docx

app = FastAPI(title="Yandex Forms — БРС Word generator")

BASE_DIR = Path(__file__).parent
CAPTURES_DIR = BASE_DIR / "captures"
OUTPUT_DIR = BASE_DIR / "output"
CAPTURES_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)


@app.get("/")
async def health() -> dict:
    """Проверка, что сервер жив (открой URL туннеля в браузере)."""
    return {"status": "ok", "message": "Yandex Forms inspector is running"}


@app.api_route("/webhook", methods=["POST", "PUT", "GET"])
@app.api_route("/", methods=["POST"])
async def capture(request: Request) -> JSONResponse:
    """Принимает запрос от Яндекс Формы и сохраняет его целиком."""
    raw_body = await request.body()
    headers = dict(request.headers)

    # Пытаемся распарсить тело как JSON для красивого вывода.
    parsed = None
    try:
        parsed = json.loads(raw_body) if raw_body else None
    except json.JSONDecodeError:
        parsed = None

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")

    # ---- Лог в консоль ----
    print("\n" + "=" * 70)
    print(f"[{timestamp}] {request.method} {request.url.path}")
    print("-- HEADERS --")
    for key, value in headers.items():
        print(f"  {key}: {value}")
    print("-- BODY (raw) --")
    print(raw_body.decode("utf-8", errors="replace") or "<пусто>")
    if parsed is not None:
        print("-- BODY (parsed JSON) --")
        print(json.dumps(parsed, ensure_ascii=False, indent=2))
    print("=" * 70 + "\n")

    # ---- Сохранение в файл ----
    capture = {
        "received_at": timestamp,
        "method": request.method,
        "path": request.url.path,
        "query": dict(request.query_params),
        "headers": headers,
        "body_raw": raw_body.decode("utf-8", errors="replace"),
        "body_parsed": parsed,
    }
    out_file = CAPTURES_DIR / f"request_{timestamp}.json"
    out_file.write_text(
        json.dumps(capture, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"[saved] {out_file}")

    # ---- Генерация Word-документа БРС ----
    doc_path = None
    if isinstance(parsed, dict) and parsed.get("answer"):
        try:
            doc_path = generate_docx(parsed, BASE_DIR, OUTPUT_DIR)
            print(f"[docx] {doc_path}")
        except Exception as exc:  # не роняем ответ Яндексу из-за ошибки генерации
            print(f"[docx ERROR] {exc!r}")

    # Яндексу важен код 200/201/202 — иначе он будет повторять запрос.
    return JSONResponse(
        status_code=200,
        content={"status": "ok", "document": doc_path.name if doc_path else None},
    )
