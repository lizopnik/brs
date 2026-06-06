from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse

from brs_generator import generate_docx

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

app = FastAPI(title="Yandex Forms — БРС Word generator")

BASE_DIR = Path(__file__).parent
CAPTURES_DIR = BASE_DIR / "captures"
OUTPUT_DIR = BASE_DIR / "output"
CAPTURES_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)


# id поля «почта», которое преподаватель вводит в форме (см. captures/)
ID_EMAIL = 126112461
# id полей ФИО и дисциплины — для текста письма
ID_FAMILIYA, ID_IMYA, ID_OTCHESTVO, ID_DISCIPLINA = 125700719, 125700737, 125700764, 125700776


@app.get("/")
async def health() -> dict:
    return {"status": "ok", "message": "Yandex Forms inspector is running"}


@app.api_route("/webhook", methods=["POST", "PUT", "GET"])
@app.api_route("/", methods=["POST"])
async def capture(request: Request) -> JSONResponse:
    raw_body = await request.body()
    headers = dict(request.headers)

    parsed = None
    try:
        parsed = json.loads(raw_body) if raw_body else None
    except json.JSONDecodeError:
        parsed = None

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")

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

    capture_data = {
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
        json.dumps(capture_data, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"[saved] {out_file}")

    doc_path = None
    if isinstance(parsed, dict) and parsed.get("answer"):
        try:
            doc_path = generate_docx(parsed, BASE_DIR, OUTPUT_DIR)
            print(f"[docx] {doc_path}")

            # Отправляем готовый документ на email, указанный в форме.
            answer_data = parsed.get("answer", {}).get("data", {})
            id_map = {}
            for entry in answer_data.values():
                if isinstance(entry, dict):
                    qid = entry.get("question", {}).get("id")
                    val = entry.get("value")
                    if qid is not None and val is not None:
                        id_map[qid] = str(val)

            email = id_map.get(ID_EMAIL, "").strip()
            familiya = id_map.get(ID_FAMILIYA, "").strip()
            imya = id_map.get(ID_IMYA, "").strip()
            otchestvo = id_map.get(ID_OTCHESTVO, "").strip()
            disciplina = id_map.get(ID_DISCIPLINA, "").strip()
            fio = " ".join(p for p in (familiya, imya, otchestvo) if p)

            if email:
                from notify import send_document
                send_document(email, doc_path, fio, disciplina)
            else:
                print("[mail] в форме нет email — письмо не отправлено")

        except Exception as exc:
            print(f"[docx ERROR] {exc!r}")

    return JSONResponse(
        status_code=200,
        content={"status": "ok", "document": doc_path.name if doc_path else None},
    )
