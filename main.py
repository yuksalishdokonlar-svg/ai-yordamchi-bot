import os
import asyncio
import re
from pathlib import Path
from threading import Thread

from flask import Flask
from ddgs import DDGS
from dotenv import load_dotenv
from google import genai
from telegram import Update
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    ContextTypes,
    filters,
)

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env", override=True)

TOKEN = os.getenv("BOT_TOKEN")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

client = None

# --- Render.com 24/7 ishlashi uchun kichik veb-server ---
web_app = Flask(__name__)


@web_app.route("/")
def home():
    return "Bot 24/7 faol ishlamoqda!"


def run_web():
    port = int(os.environ.get("PORT", 8080))
    web_app.run(host="0.0.0.0", port=port)


def keep_alive():
    t = Thread(target=run_web, daemon=True)
    t.start()


# --- Web qidiruv funksiyasi ---
def search_web(query, retries=2):
    attempt_query = query
    last_error = None

    for attempt in range(retries + 1):
        try:
            results = DDGS().text(attempt_query, region="uz-uz", max_results=5)
            if results:
                return results
        except Exception as e:
            last_error = e
            print(f"Qidiruv urinishi {attempt + 1} muvaffaqiyatsiz:", repr(e))

        words = attempt_query.split()
        if len(words) > 6:
            attempt_query = " ".join(words[:6])
        else:
            break

    if last_error:
        print("Qidiruv yakuniy xatosi:", repr(last_error))
    return []


# --- Eslatma yuborish funksiyasi ---
async def send_reminder_job(context: ContextTypes.DEFAULT_TYPE):
    job = context.job
    await context.bot.send_message(
        chat_id=job.chat_id,
        text=f"⏰ Eslatma!\n{job.data}",
    )


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Assalomu alaykum! Men sizning AI yordamchingizman.\n\n"
        "Excel, Word, soliq, buxgalteriya, yangiliklar va "
        "kunlik rejalaringiz bo‘yicha savol berishingiz mumkin.\n\n"
        "Eslatma uchun: 2 minutdan keyin yig‘ilishni boshlash"
    )


async def message_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    lower_text = text.lower()
    chat_id = update.effective_chat.id

   # --- Eslatma (Reminder) ---
    match = re.search(
        r"(\d+)\s*(soniya|sekund|minut|daqiqa|soat)\s*(?:dan\s*keyin|da|ga)?",
        lower_text,
    )

    if match and any(w in lower_text for w in ("eslat", "keyin", "taymer")):
        amount = int(match.group(1))
        unit = match.group(2)

        if unit in ("soniya", "sekund"):
            seconds = amount
            unit_name = "soniya"
        elif unit in ("minut", "daqiqa"):
            seconds = amount * 60
            unit_name = "daqiqa"
        else:
            seconds = amount * 3600
            unit_name = "soat"

        if not 5 <= seconds <= 86400:
            await update.message.reply_text("5 soniyadan 24 soatgacha vaqt kiriting.")
            return

        reminder = (text[:match.start()] + " " + text[match.end():]).strip()
        reminder = re.sub(
            r"\b(menga|iltimos|eslatib\s*qo['‘’]?y|eslat|deb)\b",
            "",
            reminder,
            flags=re.IGNORECASE,
        )
        reminder = re.sub(r"\s+", " ", reminder).strip(" ,.!-")

        if not reminder:
            await update.message.reply_text(
                "Nimani eslatishim kerakligini ham yozing. Masalan: 10 daqiqadan keyin choy ichish"
            )
            return

        await update.message.reply_text(
            f"Xo‘p! {amount} {unit_name}dan keyin «{reminder}» deb eslataman."
        )

        context.job_queue.run_once(
            send_reminder_job,
            when=seconds,
            chat_id=chat_id,
            data=reminder,
        )
        return

    # --- Web search (agar kerak bo'lsa) ---
    search_words = (
        "internetdan", "qidir", "bugun", "hozir",
        "so'nggi", "yangilik", "amaldagi",
        "2026", "soliq", "qonun",
    )

    search_results = []
    if any(word in lower_text for word in search_words):
        try:
            search_results = await asyncio.to_thread(search_web, text)
        except Exception as e:
            print("Qidiruv xatosi:", repr(e))

    search_context = "\n".join(
        f"Sarlavha: {r.get('title', '')}\n"
        f"Ma'lumot: {r.get('body', '')}\n"
        f"Havola: {r.get('href', '')}"
        for r in search_results
    )

    # --- Gemini javobi ---
    wait_msg = await update.message.reply_text("Savolingizni ko‘rib chiqyapman...")
    try:
        response = await asyncio.to_thread(
            client.models.generate_content,
            model="gemini-3.8-flash",
            contents=(
                "Sen o‘zbek tilida javob beradigan yordamchisan. "
                "Internet qidiruv natijalaridan foydalan. "
                "O‘zbekiston soliqlari va qonunlari haqida savollarda "
                "rasmiy manbalarni ustun qo‘y: soliq.uz, lex.uz, gov.uz. "
                "Topilmagan ma'lumotni o‘ylab topma. "
                "Foydalanilgan havolalarni javobda ko‘rsat.\n\n"
                "Internet qidiruv natijalari:\n"
                + (search_context or "Qidiruv bajarilmadi.")
                + "\n\nFoydalanuvchi savoli:\n"
                + text
            ),
        )

        answer = response.text or "Hozir javob tayyorlay olmadim."
        await wait_msg.delete()

        for i in range(0, len(answer), 4000):
            await update.message.reply_text(answer[i:i + 4000])

    except Exception as e:
        print("Gemini xatosi:", repr(e))
        await wait_msg.edit_text(
            "Kechirasiz, hozir AI javob bera olmadi. "
            "Birozdan keyin qayta urinib ko‘ring."
        )


def main():
    global client

    if not TOKEN:
        raise ValueError("BOT_TOKEN topilmadi. Environment Variables ni tekshiring.")

    if not GEMINI_API_KEY:
        raise ValueError("GEMINI_API_KEY topilmadi. Environment Variables ni tekshiring.")

    client = genai.Client(api_key=GEMINI_API_KEY)

    keep_alive()

    app = Application.builder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(
        MessageHandler(filters.TEXT & ~filters.COMMAND, message_handler)
    )

    print("AI yordamchi 24/7 rejimda ishga tushdi!")
    app.run_polling()


if __name__ == "__main__":
    main() 