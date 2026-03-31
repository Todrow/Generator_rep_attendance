"""
Скрипт генерации табелей явки на текущую рабочую неделю.
Генерирует 7 файлов (пн-вс), меняет дату в ячейке A3,
называет файлы по шаблону: ДеньНедели_ДД_ММ_ГГГГ.xlsx
Отправляет файлы в Telegram и на Email.
"""

import os
import shutil
import smtplib
import requests
from datetime import date, timedelta
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from openpyxl import load_workbook
import datetime

# ──────────────────────────────────────────────
# НАСТРОЙКИ — заполни перед использованием
# ──────────────────────────────────────────────

TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "")
TELEGRAM_CHAT_ID   = os.environ.get("TELEGRAM_CHAT_ID",   "")

EMAIL_SENDER   = os.environ.get("EMAIL_SENDER",   "")   # твой gmail
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD", "")   # пароль приложения Gmail
EMAIL_RECEIVER = os.environ.get("EMAIL_RECEIVER", "")   # куда слать

# TEMPLATE_PATH = "template.xlsx"   # шаблон — копия оригинального файла

# ──────────────────────────────────────────────

DAYS_RU = {
    0: "Понедельник",
    1: "Вторник",
    2: "Среда",
    3: "Четверг",
    4: "Пятница",
    5: "Суббота",
    6: "Воскресенье",
}

# Сопоставление дней недели и файлов
days_to_files = {
    'Понедельник': 'Понедельник.xlsx',
    'Вторник': 'Вторник.xlsx',
    'Среда': 'Среда.xlsx',
    'Четверг': 'Четверг.xlsx',
    'Пятница': 'Пятница.xlsx',
    'Суббота': 'Суббота.xlsx',
    'Воскресенье': None   # Нет файла
}


def get_week_dates() -> list[date]:
    """Возвращает даты текущей недели (пн–вс)."""
    today = date.today()
    monday = today - timedelta(days=today.weekday())
    return [monday + timedelta(days=i) for i in range(7)]


def generate_file() -> str:
    """Копирует шаблон, подставляет дату, возвращает путь к файлу."""
    today = datetime.datetime.now().date()
    day_name = DAYS_RU[today.weekday()]
    filename = f"{day_name} {today.strftime('%d.%m.%Y')}.xlsx"
    output_path = os.path.join("output", filename)

    os.makedirs("output", exist_ok=True)
    shutil.copy(days_to_files[day_name], output_path)

    wb = load_workbook(output_path)
    ws = wb.active

    # Меняем только дату в A3, форматируем как ДД.ММ.ГГГГ
    current_date = today.strftime("%d.%m.%Y")
    ws.cell(row=3, column=1).value = current_date

    wb.save(output_path)
    print(f"  ✓ Создан: {filename}")
    return output_path


def send_telegram(file_path: str):
    """Отправляет файлы в Telegram."""
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        print("⚠️  Telegram не настроен — пропускаем.")
        return

    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendDocument"

    # Сначала отправляем сообщение
    requests.post(
        f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage",
        data={"chat_id": TELEGRAM_CHAT_ID, "text": "📊 Табели явки на текущую неделю:"},
    )

    with open(file_path, "rb") as f:
        resp = requests.post(url, data={"chat_id": TELEGRAM_CHAT_ID}, files={"document": f})
    if resp.ok:
        print(f"  ✓ Telegram: {os.path.basename(file_path)}")
    else:
        print(f"  ✗ Telegram ошибка: {resp.text}")


def send_email(file_paths: list[str]):
    """Отправляет файлы на Email через Gmail SMTP."""
    if not EMAIL_SENDER or not EMAIL_PASSWORD or not EMAIL_RECEIVER:
        print("⚠️  Email не настроен — пропускаем.")
        return

    msg = MIMEMultipart()
    msg["From"]    = EMAIL_SENDER
    msg["To"]      = EMAIL_RECEIVER
    msg["Subject"] = f"Табели явки — неделя {date.today().strftime('%d.%m.%Y')}"
    msg.attach(MIMEText("Добрый день!\n\nВо вложении табели явки на текущую неделю.", "plain", "utf-8"))

    for path in file_paths:
        with open(path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="{os.path.basename(path)}"')
        msg.attach(part)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(EMAIL_SENDER, EMAIL_PASSWORD)
        server.sendmail(EMAIL_SENDER, EMAIL_RECEIVER, msg.as_string())

    print(f"  ✓ Email отправлен на {EMAIL_RECEIVER}")


def main():
    print("🗓  Генерация файлов...")
    week_dates = get_week_dates()
    generated = generate_file()

    print("\n📨 Отправка в Telegram...")
    send_telegram(generated)

    print("\n📧 Отправка по Email...")
    # send_email(generated)

    print("\n✅ Готово!")


if __name__ == "__main__":
    main()
