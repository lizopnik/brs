from __future__ import annotations

import json
import os
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


def _try_send_confirmation(fio: str, subject: str) -> None:
    """
    Находит преподавателя в teachers.txt по фамилии и предмету
    и отправляет ему подтверждение о получении БРС.
    """
    try:
        from notify import load_teachers, _normalize, YANDEX_EMAIL
        import smtplib
        from email.mime.multipart import MIMEMultipart
        from email.mime.text import MIMEText

        teachers = load_teachers()
        last_name = fio.split()[0] if fio else ""
        target = next(
            (t for t in teachers
             if _normalize(t.last_name) == _normalize(last_name)
             and _normalize(t.subject) == _normalize(subject)),
            None,
        )
        if target is None:
            print(f"[confirm] Преподаватель не найден в teachers.txt: {fio} / {subject}")
            return

        password = os.getenv("YANDEX_PASSWORD", "")
        if not YANDEX_EMAIL or not password:
            print("[confirm] YANDEX_EMAIL/YANDEX_PASSWORD не заданы")
            return

        body = (
            f"Уважаемый(ая) {target.fio},\n\n"
            f"Ваш БРС по предмету «{target.subject}» успешно получен и сохранен.\n\n"
            f"С уважением,\nКафедра"
        )
        msg = MIMEMultipart()
        msg["From"] = YANDEX_EMAIL
        msg["To"] = target.email
        msg["Subject"] = f"[БРС] Форма по предмету {target.subject} получена"
        msg.attach(MIMEText(body, "plain", "utf-8"))

        conn = smtplib.SMTP_SSL("smtp.yandex.ru", 465)
        conn.login(YANDEX_EMAIL, password)
        conn.sendmail(YANDEX_EMAIL, target.email, msg.as_bytes())
        conn.quit()
        print(f"[confirm] Подтверждение отправлено {target.fio} <{target.email}>")
    except Exception as e:
        print(f"[confirm] Ошибка: {e!r}")


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

            # Новое: отправляем подтверждение преподавателю
            answer_data = parsed.get("answer", {}).get("data", {})
            id_map = {}
            for entry in answer_data.values():
                if isinstance(entry, dict):
                    qid = entry.get("question", {}).get("id")
                    val = entry.get("value")
                    if qid is not None and val is not None:
                        id_map[qid] = str(val)

            familiya = id_map.get(125700719, "").strip()
            imya = id_map.get(125700737, "").strip()
            otchestvo = id_map.get(125700764, "").strip()
            disciplina = id_map.get(125700776, "").strip()
            fio = " ".join(p for p in (familiya, imya, otchestvo) if p)

            if fio and disciplina:
                _try_send_confirmation(fio, disciplina)

        except Exception as exc:
            print(f"[docx ERROR] {exc!r}")

    return JSONResponse(
        status_code=200,
        content={"status": "ok", "document": doc_path.name if doc_path else None},
    )
