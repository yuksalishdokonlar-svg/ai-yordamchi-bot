import asyncio
import io
import logging
import os
import re
import sqlite3
from datetime import datetime, timedelta
from typing import Any, Awaitable, Callable, Dict
from zoneinfo import ZoneInfo

import edge_tts
from aiohttp import web
from aiogram import BaseMiddleware, Bot, Dispatcher, F
from aiogram import BaseMiddleware, Bot, Dispatcher, F
from aiogram.filters import Command, CommandStart
from aiogram.types import BufferedInputFile, Message, TelegramObject
from docx import Document
from dotenv import load_dotenv
from google import genai
from google.genai import types

# 1. .env faylidan o'zgaruvchilarni yuklash
load_dotenv()

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
ALLOWED_USER_ID_RAW = os.getenv("ALLOWED_USER_ID")
TIMEZONE_STR = os.getenv("TIMEZONE", "Asia/Tashkent")

if not TELEGRAM_BOT_TOKEN or not GEMINI_API_KEY or not ALLOWED_USER_ID_RAW:
    raise ValueError(
        "Xatolik: .env faylida TELEGRAM_BOT_TOKEN, GEMINI_API_KEY yoki ALLOWED_USER_ID topilmadi!"
    )

ALLOWED_USER_ID = int(ALLOWED_USER_ID_RAW)
LOCAL_TZ = ZoneInfo(TIMEZONE_STR)

# 2. J.A.R.V.I.S. xarakteri
JARVIS_BASE_INSTRUCTION = """
Sen J.A.R.V.I.S.san (Just A Rather Very Intelligent System) — Toni Starkning aqlli,
vazmin, biroz kinoyali, ammo nihoyatda aniq va foydali sun'iy intellekt yordamchisi.
Qoidalar:
- Foydalanuvchiga hurmat bilan 'Janob' deb murojaat qil.
- Javoblarni qisqa, lo'nda va texnik jihatdan aniq ber.
- Ovozli javoblar uchun ortiqcha belgilar (*, #, _) ishlatma, chunki matn ovozga o'giriladi.
- Jonli ma'lumotlar (ob-havo, yangiliklar, valyuta kurslari, sport natijalari) so'ralganda Google Search natijalariga tayanib eng so'nggi faktlarni keltir.
- Hujjatlar (PDF, Word yoki matnli fayllar) yuborilganda ularning asosiy mazmuni, raqamlar va muhim xulosalarini tizimli ravishda chiqar.
- O'zbek tilida ravon javob ber.
"""

FALLBACK_MODELS = [
    "gemini-3.1-flash-lite",
    "gemini-2.5-flash-lite",
    "gemini-3.5-flash-lite",
    "gemini-3.5-flash",
]
TTS_VOICE = "uz-UZ-SardorNeural"
DB_PATH = "jarvis_memory.db"
MAX_HISTORY_MESSAGES = 20

# Eslatma rejimini aniqlash uchun kalit so'zlar
REMINDER_KEYWORDS = (
    "eslat",
    "eslatma",
    "budilnik",
    "taymer",
    "bekor qil",
    "o'chir",
    "ochir",
    "reja",
    "jadval",
)

SUPPORTED_DOC_MIMES = {
    "application/pdf": "application/pdf",
    "text/plain": "text/plain",
    "text/csv": "text/csv",
    "text/markdown": "text/markdown",
    "text/html": "text/html",
    "application/json": "application/json",
    "text/x-python": "text/plain",
    "application/javascript": "text/plain",
}

SUPPORTED_EXTENSIONS = {
    ".pdf": "application/pdf",
    ".txt": "text/plain",
    ".csv": "text/csv",
    ".md": "text/markdown",
    ".json": "text/json",
    ".py": "text/plain",
    ".js": "text/plain",
    ".html": "text/html",
    ".log": "text/plain",
}

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

client = genai.Client(api_key=GEMINI_API_KEY)
bot = Bot(token=TELEGRAM_BOT_TOKEN)
dp = Dispatcher()


# --- 3. Xavfsizlik Middleware ---
class OwnerOnlyMiddleware(BaseMiddleware):
    async def __call__(
        self,
        handler: Callable[[TelegramObject, Dict[str, Any]], Awaitable[Any]],
        event: TelegramObject,
        data: Dict[str, Any],
    ) -> Any:
        user = data.get("event_from_user")
        if user and user.id != ALLOWED_USER_ID:
            logging.warning(
                f"Ruxsatsiz kirish urinishi: ID={user.id}, Username=@{user.username}"
            )
            if isinstance(event, Message):
                await event.answer(
                    "Kechirasiz, ushbu tizim faqat avtorizatsiyadan o‘tgan egasi uchun "
                    "himoyalangan. Kirish rad etildi."
                )
            return
        return await handler(event, data)


dp.message.middleware(OwnerOnlyMiddleware())


# --- 4. SQLite Ma'lumotlar Bazasi ---
def init_db():
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS chat_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                role TEXT NOT NULL,
                text TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS user_settings (
                user_id INTEGER PRIMARY KEY,
                voice_mode INTEGER NOT NULL DEFAULT 0
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS reminders (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                remind_at TEXT NOT NULL,
                message TEXT NOT NULL,
                is_daily INTEGER NOT NULL DEFAULT 0,
                is_sent INTEGER NOT NULL DEFAULT 0
            )
            """
        )
        try:
            conn.execute(
                "ALTER TABLE reminders ADD COLUMN is_daily INTEGER NOT NULL DEFAULT 0"
            )
        except sqlite3.OperationalError:
            pass
        conn.commit()


def is_voice_mode_enabled(user_id: int) -> bool:
    with sqlite3.connect(DB_PATH) as conn:
        cursor = conn.execute(
            "SELECT voice_mode FROM user_settings WHERE user_id = ?", (user_id,)
        )
        row = cursor.fetchone()
        return bool(row[0]) if row else False


def toggle_voice_mode(user_id: int) -> bool:
    current = is_voice_mode_enabled(user_id)
    new_state = not current
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            """
            INSERT INTO user_settings (user_id, voice_mode)
            VALUES (?, ?)
            ON CONFLICT(user_id) DO UPDATE SET voice_mode = excluded.voice_mode
            """,
            (user_id, int(new_state)),
        )
        conn.commit()
    return new_state


def save_message_to_db(user_id: int, role: str, text: str):
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            "INSERT INTO chat_history (user_id, role, text) VALUES (?, ?, ?)",
            (user_id, role, text),
        )
        conn.commit()


def load_history_from_db(user_id: int) -> list[types.Content]:
    with sqlite3.connect(DB_PATH) as conn:
        cursor = conn.execute(
            """
            SELECT role, text FROM (
                SELECT id, role, text FROM chat_history
                WHERE user_id = ?
                ORDER BY id DESC
                LIMIT ?
            ) ORDER BY id ASC
            """,
            (user_id, MAX_HISTORY_MESSAGES),
        )
        rows = cursor.fetchall()

    return [
        types.Content(role=role, parts=[types.Part.from_text(text=text)])
        for role, text in rows
    ]


def clear_history_in_db(user_id: int):
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("DELETE FROM chat_history WHERE user_id = ?", (user_id,))
        conn.commit()


# --- 5. Eslatmalar (Gemini Function Calling Tools) ---
def create_reminder_tools_for_user(user_id: int):
    def set_reminder(
        remind_at_iso: str, reminder_text: str, is_daily: bool = False
    ) -> str:
        """Belgilangan sana va vaqtga bir martalik yoki har kuni takrorlanadigan eslatma o'rnatadi.

        Args:
            remind_at_iso: Birinchi eslatma vaqti ISO formatda (YYYY-MM-DD HH:MM:SS), masalan '2026-11-24 08:30:00'.
            reminder_text: Eslatma matni.
            is_daily: Har kuni shu vaqtda takrorlanishi kerak bo'lsa True, bir martalik bo'lsa False.
        """
        try:
            dt = datetime.strptime(remind_at_iso, "%Y-%m-%d %H:%M:%S").replace(
                tzinfo=LOCAL_TZ
            )
            now = datetime.now(LOCAL_TZ)

            if dt <= now:
                if is_daily:
                    while dt <= now:
                        dt += timedelta(days=1)
                else:
                    return (
                        f"Xatolik: Ko'rsatilgan vaqt ({remind_at_iso}) allaqachon o'tib ketgan. "
                        f"Hozirgi vaqt: {now.strftime('%Y-%m-%d %H:%M:%S')}."
                    )

            with sqlite3.connect(DB_PATH) as conn:
                cursor = conn.execute(
                    "INSERT INTO reminders (user_id, remind_at, message, is_daily) VALUES (?, ?, ?, ?)",
                    (
                        user_id,
                        dt.strftime("%Y-%m-%d %H:%M:%S"),
                        reminder_text,
                        int(is_daily),
                    ),
                )
                new_id = cursor.lastrowid
                conn.commit()

            mode_label = "har kuni takrorlanadigan" if is_daily else "bir martalik"
            return (
                f"Muvaffaqiyatli saqlandi [ID: {new_id}]: '{reminder_text}' ({mode_label}) "
                f"eslatmasi {dt.strftime('%Y-%m-%d %H:%M:%S')} ga o'rnatildi."
            )
        except ValueError:
            return "Xatolik: Vaqt formati noto'g'ri. 'YYYY-MM-DD HH:MM:SS' formatidan foydalanish kerak."

    def list_reminders() -> str:
        """Foydalanuvchining faol eslatmalari ro'yxatini qaytaradi."""
        with sqlite3.connect(DB_PATH) as conn:
            cursor = conn.execute(
                """
                SELECT id, remind_at, message, is_daily FROM reminders
                WHERE user_id = ? AND is_sent = 0
                ORDER BY remind_at ASC
                """,
                (user_id,),
            )
            rows = cursor.fetchall()

        if not rows:
            return "Hozirda faol eslatmalar mavjud emas."
        result = ["Faol eslatmalar ro'yxati:"]
        for r_id, r_time, r_msg, r_daily in rows:
            tag = "[Har kuni]" if r_daily else "[Bir martalik]"
            result.append(f"- ID: {r_id} | {tag} {r_time} — {r_msg}")
        return "\n".join(result)

    def delete_reminder(reminder_id: int) -> str:
        """Ko'rsatilgan ID raqami bo'yicha eslatmani bekor qiladi va bazadan o'chiradi.

        Args:
            reminder_id: O'chirilishi kerak bo'lgan eslatmaning ID raqami.
        """
        with sqlite3.connect(DB_PATH) as conn:
            cursor = conn.execute(
                "SELECT message FROM reminders WHERE id = ? AND user_id = ? AND is_sent = 0",
                (reminder_id, user_id),
            )
            row = cursor.fetchone()
            if not row:
                return f"Xatolik: ID={reminder_id} raqamli faol eslatma topilmadi."

            conn.execute(
                "DELETE FROM reminders WHERE id = ? AND user_id = ?",
                (reminder_id, user_id),
            )
            conn.commit()

        return f"Bekor qilindi: [ID: {reminder_id}] '{row[0]}' eslatmasi muvaffaqiyatli o'chirildi."

    return [set_reminder, list_reminders, delete_reminder]


def is_reminder_intent(text: str) -> bool:
    """So'rov eslatmalarga oid ekanligini tekshiradi."""
    lower_text = text.lower()
    return any(keyword in lower_text for keyword in REMINDER_KEYWORDS)


def build_chat_config(user_id: int, use_reminder_tools: bool = False):
    now_str = datetime.now(LOCAL_TZ).strftime("%Y-%m-%d %H:%M:%S (%A)")
    dynamic_instruction = (
        f"{JARVIS_BASE_INSTRUCTION}\n"
        f"Hozirgi aniq mahalliy sana va vaqt ({TIMEZONE_STR}): {now_str}. "
        f"Nisbiy vaqtlarni shu vaqt asosida hisobla."
    )

    if use_reminder_tools:
        active_tools = create_reminder_tools_for_user(user_id)
    else:
        active_tools = None

    return types.GenerateContentConfig(
        system_instruction=dynamic_instruction,
        tools=active_tools,
        temperature=0.7,
    )


async def send_message_with_fallback(
    user_id: int, message_content: Any, use_reminder_tools: bool = False
):
    """Zaxira modellar ro'yxati bo'yicha javob oladi (503/429 holatlarida avtomatik keyingi modelga o'tadi)."""
    history = load_history_from_db(user_id)
    config = build_chat_config(user_id, use_reminder_tools=use_reminder_tools)
    last_error = None

    for model_id in FALLBACK_MODELS:
        try:
            chat = client.chats.create(
                model=model_id,
                history=history,
                config=config,
            )
            response = await asyncio.to_thread(chat.send_message, message_content)
            return response
        except Exception as e:
            last_error = e
            logging.warning(
                f"{model_id} vaqtincha javob bermadi ({e}), keyingi zaxira modelga o'tilmoqda..."
            )
            await asyncio.sleep(0.5)
            continue

    raise last_error


def format_response_with_sources(response) -> str:
    """Gemini javobiga Google Search manbalari (havolalar) mavjud bo'lsa, oxiriga qo'shadi."""
    text = response.text or ""
    try:
        candidate = response.candidates[0]
        metadata = getattr(candidate, "grounding_metadata", None)
        chunks = getattr(metadata, "grounding_chunks", None) if metadata else None

        if chunks:
            sources = []
            seen_urls = set()
            for chunk in chunks:
                web = getattr(chunk, "web", None)
                if web and getattr(web, "uri", None) and web.uri not in seen_urls:
                    seen_urls.add(web.uri)
                    title = getattr(web, "title", None) or "Manba"
                    sources.append(f"• {title}: {web.uri}")
                    if len(sources) >= 3:
                        break

            if sources:
                text += "\n\nManbalar:\n" + "\n".join(sources)
    except Exception:
        pass

    return text


# --- 6. Orqa fon Eslatma Tekshiruvchisi ---
async def reminder_scheduler():
    while True:
        try:
            now = datetime.now(LOCAL_TZ)
            now_str = now.strftime("%Y-%m-%d %H:%M:%S")
            with sqlite3.connect(DB_PATH) as conn:
                cursor = conn.execute(
                    """
                    SELECT id, user_id, remind_at, message, is_daily FROM reminders
                    WHERE is_sent = 0 AND remind_at <= ?
                    """,
                    (now_str,),
                )
                due_reminders = cursor.fetchall()

                for r_id, user_id, remind_at_str, msg_text, is_daily in due_reminders:
                    prefix = "Kunlik eslatma" if is_daily else "Rejalashtirilgan eslatma"
                    alert_text = f"Janob, {prefix.lower()} vaqti keldi: {msg_text}"
                    try:
                        if is_voice_mode_enabled(user_id):
                            voice_bytes = await text_to_speech(alert_text)
                            voice_file = BufferedInputFile(
                                voice_bytes, filename="jarvis_reminder.mp3"
                            )
                            await bot.send_voice(
                                chat_id=user_id, voice=voice_file, caption=alert_text
                            )
                        else:
                            await bot.send_message(chat_id=user_id, text=f"⏰ {alert_text}")

                        if is_daily:
                            next_dt = datetime.strptime(
                                remind_at_str, "%Y-%m-%d %H:%M:%S"
                            ).replace(tzinfo=LOCAL_TZ)
                            while next_dt <= now:
                                next_dt += timedelta(days=1)

                            conn.execute(
                                "UPDATE reminders SET remind_at = ? WHERE id = ?",
                                (next_dt.strftime("%Y-%m-%d %H:%M:%S"), r_id),
                            )
                        else:
                            conn.execute(
                                "UPDATE reminders SET is_sent = 1 WHERE id = ?", (r_id,)
                            )
                        conn.commit()
                    except Exception as send_err:
                        logging.error(f"Eslatma yuborishda xatolik (ID={r_id}): {send_err}")

        except Exception as loop_err:
            logging.error(f"Scheduler siklida xatolik: {loop_err}")

        await asyncio.sleep(15)


# --- 7. Yordamchi Funksiyalar (DOCX o'qish, TTS va Javob yuborish) ---
def extract_text_from_docx(doc_bytes: bytes) -> str:
    doc = Document(io.BytesIO(doc_bytes))
    extracted_parts = []

    for para in doc.paragraphs:
        if para.text.strip():
            extracted_parts.append(para.text.strip())

    for table_idx, table in enumerate(doc.tables, start=1):
        table_rows = []
        for row in table.rows:
            cells = [cell.text.strip().replace("\n", " ") for cell in row.cells]
            table_rows.append(" | ".join(cells))
        if table_rows:
            extracted_parts.append(f"\n[Jadval #{table_idx}]\n" + "\n".join(table_rows))

    return "\n".join(extracted_parts)


def clean_text_for_tts(text: str) -> str:
    text = text.split("\n\nManbalar:\n")[0]
    text = re.sub(r"[*_`#>-]", "", text)
    return text.strip()


async def text_to_speech(text: str) -> bytes:
    clean_text = clean_text_for_tts(text)
    communicate = edge_tts.Communicate(clean_text, TTS_VOICE, rate="+5%")
    audio_buffer = io.BytesIO()
    async for chunk in communicate.stream():
        if chunk["type"] == "audio":
            audio_buffer.write(chunk["data"])
    return audio_buffer.getvalue()


async def send_long_message(message: Message, text: str):
    if len(text) > 4000:
        for i in range(0, len(text), 4000):
            await message.answer(text[i : i + 4000])
    else:
        await message.answer(text)


async def send_jarvis_reply(message: Message, reply_text: str, force_voice: bool = False):
    user_id = message.from_user.id
    use_voice = force_voice or is_voice_mode_enabled(user_id)

    if use_voice:
        await bot.send_chat_action(chat_id=message.chat.id, action="record_voice")
        voice_bytes = await text_to_speech(reply_text)
        voice_file = BufferedInputFile(voice_bytes, filename="jarvis_reply.mp3")

        if len(reply_text) <= 1000:
            await message.answer_voice(voice=voice_file, caption=reply_text)
        else:
            await message.answer_voice(voice=voice_file)
            await send_long_message(message, reply_text)
    else:
        await send_long_message(message, reply_text)


# --- 8. Telegram Handlerlar ---
@dp.message(CommandStart())
async def cmd_start(message: Message):
    user_id = message.from_user.id
    clear_history_in_db(user_id)

    welcome_text = (
        "Xush kelibsiz, Janob. J.A.R.V.I.S. tizimi to‘liq ishga tushirildi: "
        "Google Search jonli qidiruv, ovozli aloqa, eslatmalar, rasm hamda PDF/Word hujjatlar tahlili tayyor."
    )
    save_message_to_db(user_id, "model", welcome_text)
    await send_jarvis_reply(message, welcome_text, force_voice=True)


@dp.message(Command("voice"))
async def cmd_voice_toggle(message: Message):
    user_id = message.from_user.id
    new_state = toggle_voice_mode(user_id)

    if new_state:
        status_text = (
            "Ovozli javob rejimi yoqildi, Janob. Endi barcha javoblar va eslatmalar "
            "ham ovozli ko‘rinishda yuboriladi."
        )
        await send_jarvis_reply(message, status_text, force_voice=True)
    else:
        status_text = (
            "Ovozli javob rejimi o‘chirildi, Janob. Endi matnli javoblar rejimiga o‘tdik."
        )
        await message.answer(status_text)


@dp.message(F.document)
async def handle_document(message: Message):
    user_id = message.from_user.id
    doc = message.document
    file_name = doc.file_name or "nomsiz_fayl"
    ext = os.path.splitext(file_name)[1].lower()
    mime_type = doc.mime_type or ""

    is_docx = ext == ".docx" or mime_type == DOCX_MIME
    gemini_mime = SUPPORTED_DOC_MIMES.get(mime_type) or SUPPORTED_EXTENSIONS.get(ext)

    if not is_docx and not gemini_mime:
        await message.answer(
            f"Uzr so‘rayman, Janob. '{file_name}' fayl formati qo‘llab-quvvatlanmaydi. "
            "PDF, Word (.docx), TXT, CSV, MD, JSON yoki dastur kodlari (.py, .js, .html) fayllarini yuborishingiz mumkin."
        )
        return

    if doc.file_size and doc.file_size > 20 * 1024 * 1024:
        await message.answer(
            "Janob, fayl hajmi 20 MB dan katta. Telegram cheklovi sababli uni yuklab ololmayman."
        )
        return

    await bot.send_chat_action(chat_id=message.chat.id, action="typing")

    try:
        buffer = io.BytesIO()
        await bot.download(doc, destination=buffer)
        doc_bytes = buffer.getvalue()

        prompt_text = (
            message.caption
            if message.caption
            else f"'{file_name}' hujjatini diqqat bilan o‘qib chiq, asosiy mazmuni va muhim nuqtalarini tahlil qil."
        )

        if is_docx:
            extracted_text = await asyncio.to_thread(extract_text_from_docx, doc_bytes)
            if not extracted_text.strip():
                await message.answer(
                    f"Janob, '{file_name}' hujjati bo‘sh ko‘rinadi yoki undan matn ajratib bo‘lmadi."
                )
                return

            full_prompt = (
                f"Hujjat nomi: {file_name}\n"
                f"Foydalanuvchi buyrug'i: {prompt_text}\n\n"
                f"--- HUJJAT MATNI ---\n{extracted_text}"
            )
            response = await send_message_with_fallback(
                user_id, full_prompt, use_reminder_tools=False
            )
        else:
            doc_part = types.Part.from_bytes(
                data=doc_bytes,
                mime_type=gemini_mime,
            )
            response = await send_message_with_fallback(
                user_id, [doc_part, prompt_text], use_reminder_tools=False
            )

        reply_text = (
            format_response_with_sources(response)
            or "Janob, hujjatni tahlil qilishda natija olinmadi."
        )

        save_message_to_db(user_id, "user", f"[Hujjat yuborildi: {file_name}] {prompt_text}")
        save_message_to_db(user_id, "model", reply_text)

        await send_jarvis_reply(message, reply_text)

    except Exception as e:
        logging.exception("Hujjat tahlilida xatolik:")
        await message.answer(
            f"Uzr so‘rayman, Janob. Hujjatni qayta ishlash modulida nosozlik yuz berdi: {e}"
        )


@dp.message(F.voice)
async def handle_voice(message: Message):
    user_id = message.from_user.id
    await bot.send_chat_action(chat_id=message.chat.id, action="record_voice")

    try:
        buffer = io.BytesIO()
        await bot.download(message.voice, destination=buffer)
        audio_bytes = buffer.getvalue()

        audio_part = types.Part.from_bytes(
            data=audio_bytes,
            mime_type="audio/ogg",
        )

        response = await send_message_with_fallback(
            user_id,
            [
                audio_part,
                "Foydalanuvchining ovozli xabarini tingla va unga J.A.R.V.I.S. sifatida javob ber.",
            ],
            use_reminder_tools=True,
        )

        reply_text = (
            response.text
            or "Kechirasiz, Janob, ovozli xabarni tushunishda muammo yuz berdi."
        )

        save_message_to_db(user_id, "user", "[Foydalanuvchi ovozli xabar yubordi]")
        save_message_to_db(user_id, "model", reply_text)

        await send_jarvis_reply(message, reply_text, force_voice=True)

    except Exception as e:
        logging.exception("Ovozli xabarni qayta ishlashda xatolik:")
        await message.answer(
            f"Uzr so‘rayman, Janob. Audio modulida nosozlik yuz berdi: {e}"
        )


@dp.message(F.photo)
async def handle_photo(message: Message):
    user_id = message.from_user.id
    await bot.send_chat_action(chat_id=message.chat.id, action="typing")

    try:
        photo = message.photo[-1]
        buffer = io.BytesIO()
        await bot.download(photo, destination=buffer)
        image_bytes = buffer.getvalue()

        prompt_text = (
            message.caption
            if message.caption
            else "Ushbu tasvirni batafsil tahlil qil va asosiy ma'lumotlarni keltir."
        )

        image_part = types.Part.from_bytes(
            data=image_bytes,
            mime_type="image/jpeg",
        )

        response = await send_message_with_fallback(
            user_id, [image_part, prompt_text], use_reminder_tools=False
        )

        reply_text = (
            format_response_with_sources(response)
            or "Janob, tasvirni tahlil qilishda ma'lumot olinmadi."
        )

        save_message_to_db(user_id, "user", f"[Rasm yuborildi] {prompt_text}")
        save_message_to_db(user_id, "model", reply_text)

        await send_jarvis_reply(message, reply_text)

    except Exception as e:
        logging.exception("Rasm tahlilida xatolik:")
        await message.answer(
            f"Uzr so‘rayman, Janob. Vizual tahlil modulida nosozlik yuz berdi: {e}"
        )


@dp.message(F.text)
async def handle_text(message: Message):
    user_id = message.from_user.id
    await bot.send_chat_action(chat_id=message.chat.id, action="typing")

    try:
        needs_reminder = is_reminder_intent(message.text)
        response = await send_message_with_fallback(
            user_id, message.text, use_reminder_tools=needs_reminder
        )
        reply_text = (
            format_response_with_sources(response)
            or "Kechirasiz, Janob, javob shakllantirishda uzilish yuz berdi."
        )

        save_message_to_db(user_id, "user", message.text)
        save_message_to_db(user_id, "model", reply_text)

        await send_jarvis_reply(message, reply_text)

    except Exception as e:
        logging.exception("Matn so'rovida xatolik:")
        await message.answer(
            f"Uzr so‘rayman, Janob. Tizimda kutilmagan nosozlik yuz berdi: {e}"
        )


from aiohttp import web


async def health_check(request):
    return web.Response(text="J.A.R.V.I.S. 24/7 faol!")


async def start_web_server():
    app = web.Application()
    app.router.add_get("/", health_check)
    runner = web.AppRunner(app)
    await runner.setup()
    port = int(os.getenv("PORT", 8080))
    site = web.TCPSite(runner, "0.0.0.0", port)
    await site.start()


from aiohttp import web


async def health_check(request):
    return web.Response(text="J.A.R.V.I.S. 24/7 faol!")


async def start_web_server():
    app = web.Application()
    app.router.add_get("/", health_check)
    runner = web.AppRunner(app)
    await runner.setup()
    port = int(os.getenv("PORT", 8080))
    site = web.TCPSite(runner, "0.0.0.0", port)
    await site.start()


from aiohttp import web


async def health_check(request):
    return web.Response(text="J.A.R.V.I.S. 24/7 faol!")


async def start_web_server():
    app = web.Application()
    app.router.add_get("/", health_check)
    runner = web.AppRunner(app)
    await runner.setup()
    port = int(os.getenv("PORT", 8080))
    site = web.TCPSite(runner, "0.0.0.0", port)
    await site.start()


from aiohttp import web


async def health_check(request):
    return web.Response(text="J.A.R.V.I.S. 24/7 faol!")


async def start_web_server():
    app = web.Application()
    app.router.add_get("/", health_check)
    runner = web.AppRunner(app)
    await runner.setup()
    port = int(os.getenv("PORT", 8080))
    site = web.TCPSite(runner, "0.0.0.0", port)
    await site.start()


async def main():
    logging.basicConfig(level=logging.INFO)
    init_db()
    await start_web_server()
    asyncio.create_task(reminder_scheduler())
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())