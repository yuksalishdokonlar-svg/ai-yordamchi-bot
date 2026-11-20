import os
import asyncio
import logging
import re
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path

from ddgs import DDGS
from dotenv import load_dotenv
from google import genai
from google.genai import types
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

logging.basicConfig(
    filename=BASE_DIR / "bot.log",
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    encoding="utf-8",
)
logger = logging.getLogger(__name__)

client = None


class _HealthHandler(BaseHTTPRequestHandler):
    """Render platformasi 'bot ishlayaptimi' deb tekshirishi uchun oddiy javob."""

    def do_GET(self):
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"Bot ishlab turibdi.")

    def log_message(self, format, *args):
        pass  # konsolni keraksiz loglar bilan to'ldirmaslik uchun


def start_health_server():
    port = int(os.getenv("PORT", "10000"))
    server = HTTPServer(("0.0.0.0", port), _HealthHandler)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    logger.info("Health-check server %s portda ishga tushdi", port)


def search_web(query, retries=2):
    """
    DDGS ba'zan uzun/murakkab jumlalarda 'No results found' xatosini beradi,
    shuning uchun bir necha marta qayta urinamiz va zarur bo'lsa so'rovni
    qisqartiramiz.
    """
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

        # Keyingi urinishda so'rovni qisqartirib ko'ramiz (birinchi 6 so'z)
        words = attempt_query.split()
        if len(words) > 6:
            attempt_query = " ".join(words[:6])
        else:
            break

    if last_error:
        print("Qidiruv yakuniy xatosi:", repr(last_error))
    return []


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Assalomu alaykum! Men sizning AI yordamchingizman.\n\n"
        "Excel, Word, soliq, buxgalteriya, yangiliklar va "
        "kunlik rejalaringiz bo‘yicha savol berishingiz mumkin.\n\n"
        "Eslatma uchun: 2 minutdan keyin yig‘ilishni boshlash"
    )


async def message_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await process_user_text(update, context, update.message.text)


async def transcribe_voice(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Telegramdan kelgan ovozli xabarni (voice yoki audio) Gemini yordamida
    matnga aylantiradi va shu matn asosida javob qaytaradi.
    """
    voice = update.message.voice or update.message.audio

    await update.message.reply_text("Ovozli xabarni tinglayapman...")

    try:
        tg_file = await context.bot.get_file(voice.file_id)
        audio_bytes = await tg_file.download_as_bytearray()

        mime_type = "audio/ogg" if update.message.voice else "audio/mpeg"

        transcription_response = await asyncio.to_thread(
            client.models.generate_content,
            model="gemini-3.8-flash",
            contents=[
                types.Part.from_bytes(data=bytes(audio_bytes), mime_type=mime_type),
                "Ushbu audio yozuvini so'zma-so'z o'zbek tilida matnga "
                "aylantir. Faqat transkripsiya matnini qaytar, boshqa "
                "hech qanday izoh yoki qo‘shimcha so‘z yozma.",
            ],
        )

        transcribed_text = (transcription_response.text or "").strip()
        if not transcribed_text:
            await update.message.reply_text(
                "Kechirasiz, ovozli xabarni tushuna olmadim. "
                "Iltimos, yana bir bor urinib ko‘ring yoki matn yozing."
            )
            return

        await update.message.reply_text(f"🗣 Siz aytdingiz: «{transcribed_text}»")

    except Exception as e:
        print("Ovozni matnga aylantirish xatosi:", repr(e))
        await update.message.reply_text(
            "Kechirasiz, ovozli xabarni matnga aylantira olmadim. "
            "Birozdan keyin qayta urinib ko‘ring yoki matn yozing."
        )
        return

    await process_user_text(update, context, transcribed_text)


async def process_user_text(update: Update, context: ContextTypes.DEFAULT_TYPE, text: str):
    lower_text = text.lower()
    chat_id = update.effective_chat.id

    # --- Reminder (eslatma) handling ---
    match = re.search(
        r"(\d+)\s*(?:minut|daqiqa)\s*dan\s*keyin",
        lower_text,
    )

    if match:
        minutes = int(match.group(1))
        reminder = text[match.end():].strip(" ,.!")

        if not 1 <= minutes <= 1440:
            await update.message.reply_text("1 dan 1440 minutgacha kiriting.")
            return

        if not reminder:
            await update.message.reply_text(
                "Nimani eslatishim kerakligini ham yozing."
            )
            return

        await update.message.reply_text(
            f"Xo‘p! {minutes} daqiqadan keyin eslataman."
        )

        async def send_reminder():
            await asyncio.sleep(minutes * 60)
            await context.bot.send_message(
                chat_id=chat_id,
                text=f"⏰ Eslatma!\n{reminder}",
            )

        asyncio.create_task(send_reminder())
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
    await update.message.reply_text("Savolingizni ko‘rib chiqyapman...")

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

        for i in range(0, len(answer), 4000):
            await update.message.reply_text(answer[i:i + 4000])

    except Exception as e:
        print("Gemini xatosi:", repr(e))
        await update.message.reply_text(
            "Kechirasiz, hozir AI javob bera olmadi. "
            "Birozdan keyin qayta urinib ko‘ring."
        )


async def error_handler(update, context: ContextTypes.DEFAULT_TYPE):
    """
    Har qanday handler ichida kutilmagan xato chiqsa ham, dastur to'xtab
[20.11.2026 20:00] Nodirjon: qolmasligi uchun bu yerda ushlab, logga yozib qo'yamiz.
    """
    logger.error("Kutilmagan xato: %s", context.error, exc_info=context.error)
    try:
        if isinstance(update, Update) and update.effective_message:
            await update.effective_message.reply_text(
                "Kechirasiz, kutilmagan xatolik yuz berdi. "
                "Qaytadan urinib ko‘ring."
            )
    except Exception:
        pass


def build_app():
    global client

    if not TOKEN:
        raise ValueError("BOT_TOKEN topilmadi. .env faylini tekshiring.")

    if not GEMINI_API_KEY:
        raise ValueError("GEMINI_API_KEY topilmadi. .env faylini tekshiring.")

    client = genai.Client(api_key=GEMINI_API_KEY)

    app = Application.builder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(
        MessageHandler(filters.TEXT & ~filters.COMMAND, message_handler)
    )
    app.add_handler(
        MessageHandler(filters.VOICE | filters.AUDIO, transcribe_voice)
    )
    app.add_error_handler(error_handler)
    return app


def main():
    """
    MUHIM: run_polling() ni bitta jarayon ichida ikki marta ishga
    tushirish xavfli (python-telegram-bot buni to'g'ri ko'tara olmaydi
    va botni "jim" holatda qotirib qo'yishi mumkin). Shuning uchun bu
    yerda qayta urinish TSIKLI yo'q: xato chiqsa, jarayon butunlay
    to'xtaydi va uni Render platformasining o'zi avtomatik qayta
    ishga tushiradi (bu ancha ishonchli usul).
    """
    start_health_server()

    app = build_app()
    logger.info("AI yordamchi ishga tushdi!")
    print("AI yordamchi ishga tushdi! (to'xtatish uchun Ctrl+C)")
    app.run_polling()


if __name__ == "__main__":
    main()