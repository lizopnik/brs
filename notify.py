from __future__ import annotations

import argparse
import json
import os
import re
import smtplib
import time
from dataclasses import dataclass
from datetime import datetime, timedelta
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

BASE_DIR = Path(__file__).parent
CAPTURES_DIR = BASE_DIR / "captures"
OUTPUT_DIR = BASE_DIR / "output"
TEACHERS_FILE = BASE_DIR / "teachers.txt"

YANDEX_EMAIL = os.getenv("YANDEX_EMAIL", "")
YANDEX_PASSWORD = os.getenv("YANDEX_PASSWORD", "")
HEAD_EMAIL = os.getenv("HEAD_EMAIL", "")
BRS_FORM_LINK = os.getenv("BRS_FORM_LINK", "")
REPORT_TIME = os.getenv("REPORT_TIME", "18:00")
REMIND_DAY = [int(d.strip()) for d in os.getenv("REMIND_DAY", "1").split(",")]
PERIOD_DAYS = int(os.getenv("PERIOD_DAYS", "30"))

EMAIL_SUBJECT = "Сдача БРС"
EMAIL_HTML = """\
<html><body style="font-family: Arial, sans-serif; font-size:14px; color:#222;">
<p>Уважаемый(ая) <b>{name}</b>,</p>
<p>Напоминаем, что необходимо заполнить и сдать
<b>БРС</b> по дисциплине «{discipline}».</p>
<p>Для удобства используйте единое окно ввода данных:</p>
<p><a href="{form_link}">{form_link}</a></p>
<p style="color:#888; font-size:12px;">
Письмо отправлено автоматически.<br>
</p></body></html>
"""

SMTP_HOST = "smtp.yandex.ru"
SMTP_PORT = 465


@dataclass
class Teacher:
    fio: str
    subject: str
    email: str

    @property
    def last_name(self) -> str:
        return self.fio.split()[0] if self.fio else ""


def load_teachers(path: Path = TEACHERS_FILE) -> list[Teacher]:
    if not path.exists():
        raise FileNotFoundError(
            f"Файл преподавателей не найден: {path}\n"
            "Создайте teachers.txt в формате:\n"
            "Иванов Иван Иванович;Математика;ivanov@example.ru"
        )
    teachers = []
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        parts = [p.strip() for p in line.split(";")]
        if len(parts) < 3:
            continue
        fio, subject, email = parts[0], parts[1], parts[2]
        if not email or "@" not in email:
            continue
        teachers.append(Teacher(fio=fio, subject=subject, email=email))
    return teachers


def get_submitted(since: datetime) -> list[dict]:
    results = []
    if not CAPTURES_DIR.exists():
        return results

    for jfile in sorted(CAPTURES_DIR.glob("request_*.json")):
        try:
            data = json.loads(jfile.read_text(encoding="utf-8"))
        except Exception:
            continue

        ts_str = jfile.stem.replace("request_", "")
        try:
            submitted_at = datetime.strptime(ts_str[:15], "%Y%m%d_%H%M%S")
        except ValueError:
            continue

        if submitted_at < since:
            continue

        body = data.get("body_parsed") or {}
        answer_data = body.get("answer", {}).get("data", {})

        id_map: dict[int, str] = {}
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
        if fio:
            results.append({
                "fio": fio,
                "last_name": familiya,
                "subject": disciplina,
                "submitted_at": submitted_at,
            })

    return results


def _normalize(text: str) -> str:
    return re.sub(r"\s+", " ", text.strip().lower())


def find_not_submitted(
    teachers: list[Teacher],
    period_days: int = PERIOD_DAYS,
    reference_date: datetime | None = None,
) -> list[Teacher]:
    if reference_date is None:
        reference_date = datetime.now()
    since = reference_date - timedelta(days=period_days)

    submitted = get_submitted(since)
    submitted_pairs = {
        (_normalize(s["last_name"]), _normalize(s["subject"]))
        for s in submitted
    }

    return [
        t for t in teachers
        if (_normalize(t.last_name), _normalize(t.subject)) not in submitted_pairs
    ]


def _smtp_connection():
    if not YANDEX_EMAIL or not YANDEX_PASSWORD:
        raise ValueError("Не заданы YANDEX_EMAIL / YANDEX_PASSWORD в .env")
    conn = smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT)
    conn.login(YANDEX_EMAIL, YANDEX_PASSWORD)
    return conn


def send_document(to_email: str, doc_path, fio: str = "", subject: str = "") -> bool:
    """
    Отправляет сгенерированный .docx во вложении на указанный email
    (адрес, который преподаватель ввёл в форме).
    """
    to_email = (to_email or "").strip()
    if "@" not in to_email:
        print(f"[mail] некорректный адрес получателя: {to_email!r}")
        return False

    doc_path = Path(doc_path)
    if not doc_path.exists():
        print(f"[mail] файл не найден: {doc_path}")
        return False

    greeting = f"Уважаемый(ая) {fio}," if fio else "Здравствуйте,"
    disc = f" по дисциплине «{subject}»" if subject else ""
    text = (
        f"{greeting}\n\n"
        f"Во вложении — сформированный документ БРС{disc}.\n\n"
        f"Письмо отправлено автоматически."
    )

    msg = MIMEMultipart()
    msg["From"] = YANDEX_EMAIL
    msg["To"] = to_email
    msg["Subject"] = "БРС" + (f" — {subject}" if subject else "")
    msg.attach(MIMEText(text, "plain", "utf-8"))

    with open(doc_path, "rb") as fh:
        part = MIMEBase(
            "application",
            "vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        part.set_payload(fh.read())
    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition", "attachment", filename=("utf-8", "", doc_path.name)
    )
    msg.attach(part)

    try:
        conn = _smtp_connection()
        conn.sendmail(YANDEX_EMAIL, to_email, msg.as_bytes())
        conn.quit()
        print(f"[mail] документ отправлен на {to_email}")
        return True
    except Exception as e:
        print(f"[mail] ошибка отправки: {e!r}")
        return False


def send_reminders_to_all(teachers: list[Teacher], period_days: int = PERIOD_DAYS) -> None:
    not_submitted = find_not_submitted(teachers, period_days)
    if not not_submitted:
        return

    try:
        conn = _smtp_connection()
    except Exception as e:
        print(f"SMTP error: {e}")
        return

    for t in not_submitted:
        msg = MIMEMultipart()
        msg["From"] = YANDEX_EMAIL
        msg["To"] = t.email
        msg["Subject"] = EMAIL_SUBJECT
        msg.attach(MIMEText(
            EMAIL_HTML.format(name=t.fio, discipline=t.subject, form_link=BRS_FORM_LINK),
            "html", "utf-8"
        ))
        try:
            conn.sendmail(YANDEX_EMAIL, t.email, msg.as_bytes())
        except Exception:
            pass

    try:
        conn.quit()
    except Exception:
        pass


def get_new_docx_files(since_hours: int = 24) -> list[Path]:
    if not OUTPUT_DIR.exists():
        return []
    threshold = datetime.now() - timedelta(hours=since_hours)
    return sorted(
        f for f in OUTPUT_DIR.glob("*.docx")
        if datetime.fromtimestamp(f.stat().st_mtime) >= threshold
    )


def send_daily_report(since_hours: int = 24) -> bool:
    if not HEAD_EMAIL:
        return False

    new_files = get_new_docx_files(since_hours)
    now_str = datetime.now().strftime("%d.%m.%Y %H:%M")

    if new_files:
        body = (
            f"Добрый вечер!\n\n"
            f"За последние {since_hours} ч. поступило {len(new_files)} файл(ов) БРС.\n"
            f"Прикрепляю их к этому письму.\n\n"
            f"Список файлов:\n"
            + "\n".join(f"  • {f.name}" for f in new_files)
            + f"\n\nДата отчета: {now_str}"
        )
    else:
        body = (
            f"Добрый вечер!\n\n"
            f"За последние {since_hours} ч. новых файлов БРС не поступало.\n\n"
            f"Дата отчета: {now_str}"
        )

    msg = MIMEMultipart()
    msg["From"] = YANDEX_EMAIL
    msg["To"] = HEAD_EMAIL
    msg["Subject"] = f"[БРС] Ежедневный отчет {datetime.now().strftime('%d.%m.%Y')}"
    msg.attach(MIMEText(body, "plain", "utf-8"))

    for f in new_files:
        with open(f, "rb") as fh:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(fh.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", "attachment",
                        filename=("utf-8", "", f.name))
        msg.attach(part)

    try:
        conn = _smtp_connection()
        conn.sendmail(YANDEX_EMAIL, HEAD_EMAIL, msg.as_bytes())
        conn.quit()
        return True
    except Exception:
        return False


def run_daemon(teachers: list[Teacher]) -> None:
    report_h, report_m = (int(x) for x in REPORT_TIME.split(":"))
    last_report_date: datetime | None = None
    last_remind_month: int | None = None

    while True:
        now = datetime.now()

        if (
            now.hour == report_h
            and now.minute == report_m
            and (last_report_date is None or last_report_date.date() < now.date())
        ):
            send_daily_report()
            last_report_date = now

        if (
            now.day in REMIND_DAY
            and (last_remind_month is None or last_remind_month != now.month)
        ):
            send_reminders_to_all(teachers)
            last_remind_month = now.month

        time.sleep(60)


def main():
    parser = argparse.ArgumentParser(
        description="Напоминания преподавателям и отчеты заведующему."
    )
    parser.add_argument("--remind", action="store_true",
                        help="Отправить напоминания всем, кто не сдал БРС")
    parser.add_argument("--report", action="store_true",
                        help="Отправить заведующему кафедрой файлы за последние сутки")
    parser.add_argument("--daemon", action="store_true",
                        help="Запустить фоновый режим по расписанию")
    parser.add_argument("--period", type=int, default=PERIOD_DAYS,
                        help=f"Период анализа в днях (по умолч. {PERIOD_DAYS})")
    parser.add_argument("--hours", type=int, default=24,
                        help="Кол-во часов для отчета (по умолч. 24)")
    parser.add_argument("--list", action="store_true",
                        help="Показать список не сдавших без отправки писем")
    args = parser.parse_args()

    teachers = load_teachers()

    if args.list:
        not_sub = find_not_submitted(teachers, args.period)
        for t in not_sub:
            print(f"{t.fio:<35} {t.subject:<30} {t.email}")
        return

    if args.remind:
        send_reminders_to_all(teachers, args.period)

    if args.report:
        send_daily_report(args.hours)

    if args.daemon:
        run_daemon(teachers)

    if not any([args.remind, args.report, args.daemon, args.list]):
        parser.print_help()


if __name__ == "__main__":
    main()