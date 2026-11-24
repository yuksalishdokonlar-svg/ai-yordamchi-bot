import asyncio
import glob
import io
import json
import logging
import os
import re
import sqlite3
import tempfile
from datetime import datetime, timedelta
from typing import Any, Awaitable, Callable, Dict
from zoneinfo import ZoneInfo

import edge_tts
import yt_dlp
from aiohttp import web
from aiogram import BaseMiddleware, Bot, Dispatcher, F
from aiogram.filters import Command, CommandStart
from aiogram.types import BufferedInputFile, FSInputFile, Message, TelegramObject
from docx import Document
from dotenv import load_dotenv
from fpdf import FPDF
from google import genai
from google.genai import types
from openpyxl import Workbook

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
- Agar foydalanuvchi Excel (.xlsx), Word (.docx), PDF (.pdf) yoki TXT (.txt) fayl tayyorlab berishni so'rasa, HECH QACHON soxta havola yoki '[Yuklab olish: ...]' deb yozma! Buning o'rniga tegishli fayl yaratish funksiyasini (create_excel_file, create_word_file, create_pdf_file yoki create_text_file) chaqir.
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

# Fayl yaratish rejimini aniqlash uchun kalit so'zlar
FILE_GEN_KEYWORDS = (
    "excel",
    "exzel",
    "exel",
    "xlsx",
    "jadval fayl",
    "word",
    "wort",
    "docx",
    "pdf",
    "txt",
    "fayl qil",
    "fayl yarat",
    "fayl tayyorla",
    "fayl ber",
    "fayl tashla",
    "fayl jo'nat",
    "faylni ber",
    "faylni tashla",
    "hujjat tayyorla",
)

YOUTUBE_URL_PATTERN = re.compile(
    r"(https?://(?:www\.|m\.)?(?:youtube\.com/(?:watch\?v=|shorts/|embed/|live/)|youtu\.be/)[^\s]+)"
)
AUDIO_INTENT_KEYWORDS = ("audio", "mp3", "qo'shiq", "qoshiq", "musiqa", "ovoz")
LINK_ONLY_KEYWORDS = ("silka", "havola", "link", "url")

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


# --- 5. Eslatmalar va Fayl Yaratish (Gemini Function Calling Tools) ---
def _safe_filename(filename: str, default_ext: str) -> str:
    clean = re.sub(r'[\\/*?:"<>|]', "_", filename.strip())
    if not clean:
        clean = f"jarvis_hujjat{default_ext}"
    if not clean.lower().endswith(default_ext):
        clean += default_ext
    return clean


def create_tools_for_user(user_id: int, generated_files: list):
    def set_reminder(
        remind_at_iso: str, reminder_text: str, is_daily: bool = False
    ) -> str:
        """Belgilangan sana va vaqtga bir martalik yoki har kuni takrorlanadigan eslatma o'rnatadi."""
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
        """Ko'rsatilgan ID raqami bo'yicha eslatmani bekor qiladi va bazadan o'chiradi."""
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

    def create_excel_file(filename: str, sheet_title: str, rows_json: str) -> str:
        """Haqiqiy Excel (.xlsx) fayl yaratadi va foydalanuvchiga yuborish uchun tayyorlaydi."""
        try:
            rows = json.loads(rows_json)
            wb = Workbook()
            ws = wb.active
            ws.title = (sheet_title or "Sheet1")[:31]

            for row in rows:
                if isinstance(row, list):
                    ws.append(row)
                else:
                    ws.append([str(row)])

            buf = io.BytesIO()
            wb.save(buf)
            safe_name = _safe_filename(filename, ".xlsx")
            generated_files.append((safe_name, buf.getvalue()))
            return f"Excel fayl '{safe_name}' muvaffaqiyatli yaratildi va yuborish navbatiga qo'shildi."
        except Exception as e:
            return f"Excel fayl yaratishda xatolik: {e}"

    def create_word_file(filename: str, title: str, content: str) -> str:
        """Haqiqiy Word (.docx) hujjat yaratadi va foydalanuvchiga yuborish uchun tayyorlaydi."""
        try:
            doc = Document()
            if title:
                doc.add_heading(title, level=1)
            for line in content.split("\n"):
                if line.strip():
                    doc.add_paragraph(line.strip())

            buf = io.BytesIO()
            doc.save(buf)
            safe_name = _safe_filename(filename, ".docx")
            generated_files.append((safe_name, buf.getvalue()))
            return f"Word fayl '{safe_name}' muvaffaqiyatli yaratildi va yuborish navbatiga qo'shildi."
        except Exception as e:
            return f"Word fayl yaratishda xatolik: {e}"

    def create_pdf_file(filename: str, title: str, content: str) -> str:
        """Haqiqiy PDF (.pdf) hujjat yaratadi va foydalanuvchiga yuborish uchun tayyorlaydi."""
        try:
            replacements = {
                "‘": "'",
                "’": "'",
                "“": '"',
                "”": '"',
                "—": "-",
                "–": "-",
                "…": "...",
            }
            clean_title = title or ""
            clean_content = content or ""
            for k, v in replacements.items():
                clean_title = clean_title.replace(k, v)
                clean_content = clean_content.replace(k, v)

            pdf = FPDF()
            pdf.add_page()
            pdf.set_auto_page_break(auto=True, margin=15)

            if clean_title:
                pdf.set_font("Helvetica", style="B", size=16)
                safe_title = clean_title.encode("latin-1", "replace").decode("latin-1")
                pdf.multi_cell(0, 10, text=safe_title)
                pdf.ln(4)

            pdf.set_font("Helvetica", size=12)
            safe_body = clean_content.encode("latin-1", "replace").decode("latin-1")
            pdf.multi_cell(0, 8, text=safe_body)

            pdf_bytes = bytes(pdf.output())
            safe_name = _safe_filename(filename, ".pdf")
            generated_files.append((safe_name, pdf_bytes))
            return f"PDF fayl '{safe_name}' muvaffaqiyatli yaratildi va yuborish navbatiga qo'shildi."
        except Exception as e:
            return f"PDF fayl yaratishda xatolik: {e}"

    def create_text_file(filename: str, content: str) -> str:
        """Haqiqiy TXT (.txt) matn fayli yaratadi va foydalanuvchiga yuborish uchun tayyorlaydi."""
        try:
            safe_name = _safe_filename(filename, ".txt")
            generated_files.append((safe_name, content.encode("utf-8")))
            return f"TXT fayl '{safe_name}' muvaffaqiyatli yaratildi va yuborish navbatiga qo'shildi."
        except Exception as e:
            return f"TXT fayl yaratishda xatolik: {e}"

    return [
        set_reminder,
        list_reminders,
        delete_reminder,
        create_excel_file,
        create_word_file,
        create_pdf_file,
        create_text_file,
    ]


def should_enable_tools(text: str) -> bool:
    """So'rov eslatma yoki fayl yaratishga oid ekanligini tekshiradi."""
    lower_text = text.lower()
    return any(k in lower_text for k in REMINDER_KEYWORDS) or any(
        k in lower_text for k in FILE_GEN_KEYWORDS
    )


def build_chat_config(user_id: int, tools_list: list | None = None):
    now_str = datetime.now(LOCAL_TZ).strftime("%Y-%m-%d %H:%M:%S (%A)")
    dynamic_instruction = (
        f"{JARVIS_BASE_INSTRUCTION}\n"
        f"Hozirgi aniq mahalliy sana va vaqt ({TIMEZONE_STR}): {now_str}. "
        f"Nisbiy vaqtlarni shu vaqt asosida hisobla."
    )

    return types.GenerateContentConfig(
        system_instruction=dynamic_instruction,
        tools=tools_list if tools_list else None,
        temperature=0.7,
    )


async def send_message_with_fallback(
    user_id: int, message_content: Any, use_tools: bool = False
):
    """Zaxira modellar ro'yxati bo'yicha javob va yaratilgan fayllarni oladi."""
    history = load_history_from_db(user_id)
    last_error = None

    for model_id in FALLBACK_MODELS:
        generated_files: list[tuple[str, bytes]] = []
        tools_list = (
            create_tools_for_user(user_id, generated_files) if use_tools else None
        )
        config = build_chat_config(user_id, tools_list=tools_list)

        try:
            chat = client.chats.create(
                model=model_id,
                history=history,
                config=config,
            )
            response = await asyncio.to_thread(chat.send_message, message_content)
            return response, generated_files
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


# --- 7. Yordamchi Funksiyalar (DOCX o'qish, TTS, YouTube va Javob/Fayl yuborish) ---
def _prepare_youtube_cookie_file(out_dir: str) -> str | None:
    """YOUTUBE_COOKIES matnidagi probellarni Netscape Tab (\t) formatiga tuzatib fayl yaratadi."""
    raw_cookies = os.getenv("YOUTUBE_COOKIES", "").strip()
    if not raw_cookies:
        return None

    fixed_lines = ["# Netscape HTTP Cookie File"]
    for raw_line in raw_cookies.replace("\r\n", "\n").split("\n"):
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue

        if "\t" in line:
            parts = line.split("\t")
        else:
            # Telefondan nusxalanganda Tab o'rniga probel tushib qolgan holatni tuzatish
            parts = line.split(None, 6)

        if len(parts) == 7:
            fixed_lines.append("\t".join(parts))

    if len(fixed_lines) <= 1:
        return None

    cookie_path = os.path.join(out_dir, "youtube_cookies.txt")
    with open(cookie_path, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(fixed_lines) + "\n")
    return cookie_path


def download_youtube_media(
    query_or_url: str, audio_only: bool, out_dir: str, link_only: bool = False
) -> tuple[str | None, str, str]:
    """YouTube havolasi yoki qidiruv so'zi bo'yicha media va haqiqiy havola oladi."""
    out_tmpl = os.path.join(out_dir, "%(title).60s.%(ext)s")
    is_direct_url = bool(YOUTUBE_URL_PATTERN.search(query_or_url))
    target = query_or_url if is_direct_url else f"ytsearch1:{query_or_url}"
    cookie_file = _prepare_youtube_cookie_file(out_dir)

    base_opts = {
        "noplaylist": True,
        "default_search": "ytsearch",
        "quiet": True,
        "no_warnings": True,
    }
    if cookie_file:
        base_opts["cookiefile"] = cookie_file

    try:
        # 1-bosqich: Video sahifasini ochmasdan tezda sarlavha va haqiqiy YouTube havolasini olamiz
        flat_opts = {**base_opts, "extract_flat": "in_playlist", "skip_download": True}
        with yt_dlp.YoutubeDL(flat_opts) as ydl:
            info = ydl.extract_info(target, download=False)
            if isinstance(info, dict) and "entries" in info and info["entries"]:
                info = info["entries"][0]

            if not isinstance(info, dict):
                raise RuntimeError("YouTube qidiruvidan natija qaytmadi.")

            title = info.get("title") or "YouTube Media"
            video_id = info.get("id")
            raw_url = info.get("url") or info.get("webpage_url") or ""
            if raw_url.startswith("http"):
                video_url = raw_url
            elif video_id:
                video_url = f"https://www.youtube.com/watch?v={video_id}"
            else:
                video_url = query_or_url if is_direct_url else ""

        if link_only:
            return None, title, video_url

        # 2-bosqich: Media faylni yuklab olish (agar bloklansa ham video_url saqlanib qoladi)
        if audio_only:
            dl_opts = {**base_opts, "format": "bestaudio/best", "outtmpl": out_tmpl}
        else:
            dl_opts = {
                **base_opts,
                "format": "best[ext=mp4][height>=360]/best[height>=360]/best",
                "outtmpl": out_tmpl,
            }

        try:
            with yt_dlp.YoutubeDL(dl_opts) as ydl:
                dl_info = ydl.extract_info(video_url or target, download=True)
                if isinstance(dl_info, dict) and "entries" in dl_info and dl_info["entries"]:
                    dl_info = dl_info["entries"][0]

                title = dl_info.get("title") or title
                downloaded_path = ydl.prepare_filename(dl_info)
                if not os.path.exists(downloaded_path):
                    files = [
                        p
                        for p in glob.glob(os.path.join(out_dir, "*"))
                        if os.path.abspath(p) != os.path.abspath(cookie_file or "")
                    ]
                    if not files:
                        return None, title, video_url
                    downloaded_path = files[0]

                return downloaded_path, title, video_url
        except Exception as dl_err:
            logging.warning(f"Media fayl yuklanmadi, havola qaytariladi: {dl_err}")
            return None, title, video_url
    finally:
        if cookie_file and os.path.exists(cookie_file):
            try:
                os.remove(cookie_file)
            except OSError:
                pass


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


async def send_generated_files(message: Message, generated_files: list[tuple[str, bytes]]):
    for file_name, file_bytes in generated_files:
        doc_file = BufferedInputFile(file_bytes, filename=file_name)
        await message.answer_document(document=doc_file)


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
        "ovozli aloqa, eslatmalar, YouTube video/qo‘shiq yuklash, rasm hamda Excel, Word, PDF va TXT hujjatlarni tahlil qilish va yaratish tayyor."
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
            response, generated_files = await send_message_with_fallback(
                user_id, full_prompt, use_tools=True
            )
        else:
            doc_part = types.Part.from_bytes(
                data=doc_bytes,
                mime_type=gemini_mime,
            )
            response, generated_files = await send_message_with_fallback(
                user_id, [doc_part, prompt_text], use_tools=True
            )

        reply_text = (
            format_response_with_sources(response)
            or "Janob, hujjatni tahlil qilishda natija olinmadi."
        )

        save_message_to_db(user_id, "user", f"[Hujjat yuborildi: {file_name}] {prompt_text}")
        save_message_to_db(user_id, "model", reply_text)

        await send_jarvis_reply(message, reply_text)
        if generated_files:
            await send_generated_files(message, generated_files)

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

        response, generated_files = await send_message_with_fallback(
            user_id,
            [
                audio_part,
                "Foydalanuvchining ovozli xabarini tingla va unga J.A.R.V.I.S. sifatida javob ber. Agar fayl yaratish so'ralgan bo'lsa, tegishli fayl yaratish funksiyasidan foydalan.",
            ],
            use_tools=True,
        )

        reply_text = (
            response.text
            or "Kechirasiz, Janob, ovozli xabarni tushunishda muammo yuz berdi."
        )

        save_message_to_db(user_id, "user", "[Foydalanuvchi ovozli xabar yubordi]")
        save_message_to_db(user_id, "model", reply_text)

        await send_jarvis_reply(message, reply_text, force_voice=True)
        if generated_files:
            await send_generated_files(message, generated_files)

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

        response, generated_files = await send_message_with_fallback(
            user_id, [image_part, prompt_text], use_tools=True
        )

        reply_text = (
            format_response_with_sources(response)
            or "Janob, tasvirni tahlil qilishda ma'lumot olinmadi."
        )

        save_message_to_db(user_id, "user", f"[Rasm yuborildi] {prompt_text}")
        save_message_to_db(user_id, "model", reply_text)

        await send_jarvis_reply(message, reply_text)
        if generated_files:
            await send_generated_files(message, generated_files)

    except Exception as e:
        logging.exception("Rasm tahlilida xatolik:")
        await message.answer(
            f"Uzr so‘rayman, Janob. Vizual tahlil modulida nosozlik yuz berdi: {e}"
        )


@dp.message(F.text)
async def handle_text(message: Message):
    user_id = message.from_user.id
    text_body = message.text or ""

    # 1. YouTube havolasi yoki nomi bo'yicha video/qo'shiq qidirib yuklash
    yt_match = YOUTUBE_URL_PATTERN.search(text_body)
    lower_text = text_body.lower()
    has_yt_keyword = any(
        w in lower_text for w in ("youtube", "yutub", "yutube")
    ) and any(
        w in lower_text
        for w in (
            "top",
            "yukl",
            "tashla",
            "ber",
            "qo'y",
            "qoy",
            "mp3",
            "audio",
            "video",
            "qo'shiq",
            "qoshiq",
            "musiqa",
            "silka",
            "havola",
            "link",
        )
    )

    if yt_match or has_yt_keyword:
        if yt_match:
            yt_target = yt_match.group(1)
        else:
            clean_query = re.sub(
                r"\b(youtube|yutub|yutube|dan|ni|bironta|bir|video|videosini|qo'shiq|qoshiq|qo'shig'ini|qoshig'ini|qoʻshigʻini|musiqa|audio|mp3|topib|top|yuklab|yuklaver|yukla|tashlab|tashla|ber|bersang|bersinda|qilib|shuni|silkasini|silka|havolasini|havola|linkini|link|ha|ham)\b",
                " ",
                lower_text,
            )
            clean_query = re.sub(r"\s+", " ", clean_query).strip()
            yt_target = clean_query or text_body

        audio_only = any(k in lower_text for k in AUDIO_INTENT_KEYWORDS)
        link_only = any(k in lower_text for k in LINK_ONLY_KEYWORDS) and not any(
            k in lower_text for k in ("yukl", "mp3", "audio", "video", "fayl")
        )

        status_msg = await message.answer(
            f"Janob, '{yt_target}' bo‘yicha YouTube havolasi qidirilmoqda..."
            if link_only
            else (
                f"Janob, '{yt_target}' bo‘yicha audio va havola qidirilmoqda, biroz kuting..."
                if audio_only
                else f"Janob, '{yt_target}' bo‘yicha video va havola qidirilmoqda, biroz kuting..."
            )
        )
        await bot.send_chat_action(
            chat_id=message.chat.id,
            action="typing" if link_only else ("upload_voice" if audio_only else "upload_video"),
        )

        try:
            with tempfile.TemporaryDirectory() as tmp_dir:
                file_path, title, video_url = await asyncio.to_thread(
                    download_youtube_media, yt_target, audio_only, tmp_dir, link_only
                )

                if link_only or not file_path:
                    await status_msg.edit_text(
                        f"🎬 {title}\n🔗 Havola: {video_url}"
                    )
                    save_message_to_db(user_id, "user", text_body)
                    save_message_to_db(
                        user_id,
                        "model",
                        f"YouTube’dan '{title}' havolasi topildi: {video_url}",
                    )
                    return

                file_size_mb = os.path.getsize(file_path) / (1024 * 1024)
                url_part = f"\n🔗 Havola: {video_url}" if video_url else ""

                if file_size_mb > 49.5:
                    await status_msg.edit_text(
                        f"Janob, '{title}' fayli hajmi ({file_size_mb:.1f} MB) Telegram botining 50 MB cheklovidan katta, "
                        f"shuning uchun faylni jo‘nata olmadim. Quyidagi havola orqali ko‘rishingiz mumkin:{url_part}"
                    )
                    return

                media_file = FSInputFile(file_path)
                if audio_only:
                    await message.answer_audio(
                        audio=media_file,
                        title=title,
                        caption=f"🎵 {title}{url_part}",
                    )
                else:
                    await message.answer_video(
                        video=media_file,
                        caption=f"🎬 {title}{url_part}",
                    )

                await status_msg.delete()
                save_message_to_db(user_id, "user", text_body)
                save_message_to_db(
                    user_id,
                    "model",
                    f"YouTube’dan '{title}' ({'audio' if audio_only else 'video'}) va havolasi yuborildi: {video_url}",
                )
        except Exception as e:
            logging.exception("YouTube yuklashda xatolik:")
            await status_msg.edit_text(
                f"Uzr so‘rayman, Janob. {e}"
            )
        return

    # 2. Oddiy matnli suhbat, eslatmalar va hujjat yaratish
    await bot.send_chat_action(chat_id=message.chat.id, action="typing")

    try:
        needs_tools = should_enable_tools(text_body)
        response, generated_files = await send_message_with_fallback(
            user_id, text_body, use_tools=needs_tools
        )
        reply_text = (
            format_response_with_sources(response)
            or "Kechirasiz, Janob, javob shakllantirishda uzilish yuz berdi."
        )

        save_message_to_db(user_id, "user", text_body)
        save_message_to_db(user_id, "model", reply_text)

        await send_jarvis_reply(message, reply_text)
        if generated_files:
            await send_generated_files(message, generated_files)

    except Exception as e:
        logging.exception("Matn so'rovida xatolik:")
        await message.answer(
            f"Uzr so‘rayman, Janob. Tizimda kutilmagan nosozlik yuz berdi: {e}"
        )


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