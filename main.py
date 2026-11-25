import asyncio
import glob
import html
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

import aiohttp
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
- Jonli ma'lumotlar (bugungi ob-havo, yangiliklar, valyuta kurslari, sport natijalari, hozirgi soat/sana holatidagi voqealar) so'ralganda Google Search natijalariga tayanib eng so'nggi faktlarni keltir.
- Agar foydalanuvchi ochiq Telegram kanallari (t.me/... yoki @username) yoki Telegram'dagi ma'lumotlarni topishni so'rasa, Google Search (site:t.me) natijalaridan foydalanib haqiqiy ochiq Telegram kanallari va postlarining t.me havolalarini ro'yxat qilib ber.
- QAT'IY XAVFSIZLIK QOIDASI: O'zbekiston Respublikasi qonunchiligida taqiqlangan ekstremistik, terroristik, radikal va aqidaparastlik g'oyalarini targ'ib qiluvchi ma'ruzalar, tashkilotlar, kanallar va videolarni HECH QACHON qidirmagin, topmagin, tavsiya qilmagin va yuklab bermagin.
- Hujjatlar (PDF, Word yoki matnli fayllar) yuborilganda ularning asosiy mazmuni, raqamlar va muhim xulosalarini tizimli ravishda chiqar.
- Agar foydalanuvchi Excel (.xlsx), Word (.docx), PDF (.pdf) yoki TXT (.txt) fayl tayyorlab berishni so'rasa, HECH QACHON soxta havola yoki '[Yuklab olish: ...]' deb yozma! Buning o'rniga tegishli fayl yaratish funksiyasini (create_excel_file, create_word_file, create_pdf_file yoki create_text_file) chaqir.
- Agar foydalanuvchi o'zining YouTube akkauntidagi videolarni (keyinroq ko'rish, yoqtirilganlar, obunalar, tarix) ko'rishni yoki yuklab olishni so'rasa, tegishli YouTube akkaunt funksiyalaridan (get_my_youtube_videos yoki download_from_my_youtube) foydalan.
- O'zbek tilida ravon javob ber.
"""

FALLBACK_MODELS = [
    "gemini-3.1-flash-lite",
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
    "hujjat tayyorla",
)

# Shaxsiy YouTube akkaunt ma'lumotlari kalit so'zlari
MY_YT_KEYWORDS = (
    "akaunt",
    "akkaunt",
    "profil",
    "keyinroq",
    "watch later",
    "yoqtirgan",
    "yoqtirilgan",
    "like",
    "obuna",
    "subscription",
    "tarix",
    "history",
    "pleylist",
    "playlist",
    "kanalim",
    "saqlangan",
)

# Ochiq Telegram qidiruv kalit so'zlari
TELEGRAM_SEARCH_KEYWORDS = (
    "telegram",
    "t.me/",
    "telegramdan",
    "telegramdagi",
    "kanaldan",
    "kanallardan",
)

# O'zbekiston Respublikasida taqiqlangan ekstremistik/terroristik tashkilotlar, oqimlar va ma'ruzalar filtri
UZB_PROHIBITED_KEYWORDS = (
    "hizb ut-tahrir",
    "hizb ut tahrir",
    "hizbut tahrir",
    "hizbut-tahrir",
    "akromiylar",
    "akromiya",
    "vahhobiylik",
    "vahobiy",
    "jihodchilar",
    "jihod",
    "jihad",
    "katibat",
    "tavhid va jihod",
    "tavhid va jihad",
    "islomiy davlat",
    "daish",
    "daesh",
    "isis",
    "isil",
    "al-qoida",
    "al qoida",
    "al-qaeda",
    "turkiston islom",
    "o'zbekiston islomiy harakati",
    "ozbekiston islomiy harakati",
    "imom buxoriy katibasi",
    "jabhat an-nusra",
    "nusra front",
    "nurcular",
    "nurchilar",
    "tablig'i jamoat",
    "tablighi jamaat",
    "abu saloh",
    "abu salox",
    "abdulloh zufar",
    "sodiq samarqandiy",
    "mahmud abdulmo'min",
    "mahmud abdulmomin",
    "rashod qori",
    "obid qori",
    "tohir yo'ldosh",
    "tohir yoldosh",
    "juma namanganiy",
    "ekstremiz",
    "extremis",
    "terror",
    "aqidaparast",
    "takfir",
)

YOUTUBE_URL_PATTERN = re.compile(
    r"(https?://(?:www\.|m\.)?(?:youtube\.com/(?:watch\?v=|shorts/|embed/|live/|playlist\?list=|@)|youtu\.be/)[^\s]+)"
)
TELEGRAM_URL_PATTERN = re.compile(
    r"(?:https?://)?(?:t\.me|telegram\.me)/(?:s/)?([a-zA-Z0-9_]{4,})(?:/(\d+))?"
)
YT_ACTION_KEYWORDS = (
    "top",
    "yukl",
    "tort",
    "skachat",
    "tashla",
    "ber",
    "qo'y",
    "qoy",
    "mp3",
    "mp4",
    "audio",
    "video",
    "qo'shiq",
    "qoshiq",
    "qoʻshiq",
    "to'plam",
    "toplam",
    "musiqa",
    "klip",
    "konsert",
    "silka",
    "havola",
    "link",
    "varyant",
    "variant",
)
AUDIO_INTENT_KEYWORDS = (
    "audio",
    "mp3",
    "qo'shiq",
    "qoshiq",
    "qoʻshiq",
    "to'plam",
    "toplam",
    "musiqa",
    "ovoz",
)
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


# --- 5. Taqiqlangan Ma'ruzalar / Kanallar Filtri, Ochiq Telegram va YouTube Funksiyalari ---
def _get_custom_blocked_channels() -> tuple[str, ...]:
    raw = os.getenv("BLOCKED_YT_CHANNELS", "")
    return tuple(c.strip().lower() for c in raw.split(",") if c.strip())


def is_prohibited_uzb_content(
    text: str, channel: str = "", description: str = ""
) -> bool:
    """O'zbekiston Respublikasida taqiqlangan ma'ruzalar, oqimlar va bloklangan kanallarni aniqlaydi."""
    normalized = f"{text} {channel} {description}".lower()
    normalized = normalized.replace("ʻ", "'").replace("’", "'").replace("‘", "'")

    if any(bad in normalized for bad in UZB_PROHIBITED_KEYWORDS):
        return True

    custom_blocked = _get_custom_blocked_channels()
    if custom_blocked and any(ch in normalized for ch in custom_blocked):
        return True

    return False


async def fetch_public_telegram_channel_posts(
    channel_or_url: str, keyword: str = "", limit: int = 6
) -> list[dict]:
    """Ochiq Telegram kanalining (t.me/s/<username>) oxirgi postlarini o'qiydi."""
    if is_prohibited_uzb_content(f"{channel_or_url} {keyword}"):
        raise PermissionError(
            "O‘zbekiston Respublikasi qonunchiligida taqiqlangan manbalar tekshirilmaydi."
        )

    m = TELEGRAM_URL_PATTERN.search(channel_or_url)
    if m:
        username = m.group(1)
    else:
        username = re.sub(r"[^a-zA-Z0-9_]", "", channel_or_url.lstrip("@").strip())

    if not username or len(username) < 4:
        return []

    web_url = f"https://t.me/s/{username}"
    if keyword:
        web_url += f"?q={keyword}"

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    }
    async with aiohttp.ClientSession(headers=headers) as session:
        async with session.get(web_url, timeout=aiohttp.ClientTimeout(total=15)) as resp:
            if resp.status != 200:
                return []
            page_html = await resp.text()

    blocks = re.findall(
        r'<div class="tgme_widget_message_wrap[\s\S]*?(?=<div class="tgme_widget_message_wrap|</section>)',
        page_html,
    )
    posts = []
    kw_lower = keyword.lower().strip()

    for block in reversed(blocks):
        post_link_m = re.search(
            rf'href="(https://t\.me/{re.escape(username)}/\d+)"', block
        )
        post_url = post_link_m.group(1) if post_link_m else f"https://t.me/{username}"

        text_m = re.search(
            r'<div class="tgme_widget_message_text[^"]*"[^>]*>([\s\S]*?)</div>', block
        )
        raw_text = text_m.group(1) if text_m else ""
        clean_text = re.sub(r"<br\s*/?>", "\n", raw_text)
        clean_text = re.sub(r"<[^>]+>", "", clean_text)
        clean_text = html.unescape(clean_text).strip()

        media_tag = ""
        if "tgme_widget_message_audio" in block or "tgme_widget_message_document" in block:
            doc_m = re.search(
                r'<div class="tgme_widget_message_document_title"[^>]*>([\s\S]*?)</div>',
                block,
            )
            doc_name = (
                html.unescape(re.sub(r"<[^>]+>", "", doc_m.group(1))).strip()
                if doc_m
                else "Audio/Fayl"
            )
            media_tag = f"[Fayl: {doc_name}] "
        elif "tgme_widget_message_video" in block:
            media_tag = "[Video] "

        combined = f"{media_tag}{clean_text}".strip()
        if not combined:
            continue
        if is_prohibited_uzb_content(combined, username):
            continue
        if kw_lower and kw_lower not in combined.lower():
            continue

        posts.append(
            {
                "channel": f"@{username}",
                "text": combined[:350],
                "url": post_url,
            }
        )
        if len(posts) >= limit:
            break

    return posts


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
            parts = line.split(None, 6)

        if len(parts) == 7:
            fixed_lines.append("\t".join(parts))

    if len(fixed_lines) <= 1:
        return None

    cookie_path = os.path.join(out_dir, "youtube_cookies.txt")
    with open(cookie_path, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(fixed_lines) + "\n")
    return cookie_path


def _resolve_yt_account_feed(section: str) -> tuple[str, str]:
    """Bo'lim nomini yt-dlp maxsus feed identifikatoriga aylantiradi."""
    s = (section or "").lower().strip()
    if any(k in s for k in ("like", "yoqtir")):
        return ":ytliked", "Yoqtirilgan videolar (Liked Videos)"
    if any(k in s for k in ("sub", "obuna")):
        return ":ytsubs", "Obunalar lentasi (Subscriptions)"
    if any(k in s for k in ("hist", "tarix", "ko'rgan", "korgan")):
        return ":ythistory", "Ko'rish tarixi (History)"
    if any(k in s for k in ("rec", "tavsiya", "bosh sahifa", "home")):
        return ":ytrec", "Tavsiya etilgan videolar"
    return ":ytwatchlater", "Keyinroq ko'rish (Watch Later)"


def fetch_my_youtube_entries(
    section: str = "watch_later", limit: int = 10
) -> tuple[str, list[dict]]:
    """Foydalanuvchining shaxsiy YouTube akkauntidagi videolar ro'yxatini qaytaradi."""
    feed_target, feed_label = _resolve_yt_account_feed(section)
    limit = max(1, min(int(limit), 25))

    with tempfile.TemporaryDirectory() as tmp_dir:
        cookie_file = _prepare_youtube_cookie_file(tmp_dir)
        if not cookie_file:
            raise RuntimeError(
                "YOUTUBE_COOKIES topilmadi. Shaxsiy akkaunt ma'lumotlari uchun Render’da YOUTUBE_COOKIES sozlangan bo‘lishi kerak."
            )

        opts = {
            "cookiefile": cookie_file,
            "extract_flat": "in_playlist",
            "playlistend": limit + 5,
            "skip_download": True,
            "quiet": True,
            "no_warnings": True,
        }

        with yt_dlp.YoutubeDL(opts) as ydl:
            info = ydl.extract_info(feed_target, download=False)

        entries = info.get("entries") if isinstance(info, dict) else None
        if not entries:
            return feed_label, []

        results = []
        for entry in entries:
            if not isinstance(entry, dict):
                continue
            title = entry.get("title") or "Nomsiz video"
            uploader = entry.get("uploader") or entry.get("channel") or ""
            desc = entry.get("description") or ""

            if is_prohibited_uzb_content(title, uploader, desc):
                continue

            vid_id = entry.get("id") or ""
            raw_url = entry.get("url") or entry.get("webpage_url") or ""
            if raw_url.startswith("http"):
                url = raw_url
            elif vid_id:
                url = f"https://www.youtube.com/watch?v={vid_id}"
            else:
                url = ""

            results.append(
                {
                    "index": len(results) + 1,
                    "title": title,
                    "uploader": uploader,
                    "url": url,
                }
            )
            if len(results) >= limit:
                break

        return feed_label, results


def download_youtube_media(
    query_or_url: str, audio_only: bool, out_dir: str, link_only: bool = False
) -> tuple[str | None, str, str, list[dict]]:
    """YouTube havolasi yoki qidiruv so'zi bo'yicha xavfsiz media, asosiy havola va variantlar ro'yxatini qaytaradi."""
    if is_prohibited_uzb_content(query_or_url):
        raise PermissionError(
            "O‘zbekiston Respublikasi qonunchiligida taqiqlangan ma’ruzalar, ekstremistik oqimlar "
            "yoki bloklangan kanal va videolarni qidira olmayman hamda yuklab bera olmayman."
        )

    out_tmpl = os.path.join(out_dir, "%(title).60s.%(ext)s")
    is_direct_url = bool(YOUTUBE_URL_PATTERN.search(query_or_url))
    target = query_or_url if is_direct_url else f"ytsearch8:{query_or_url}"
    cookie_file = _prepare_youtube_cookie_file(out_dir)

    base_opts = {
        "noplaylist": True,
        "default_search": "ytsearch",
        "quiet": True,
        "no_warnings": True,
        "extractor_args": {"youtube": {"player_client": ["android", "web", "mweb"]}},
    }
    if cookie_file:
        base_opts["cookiefile"] = cookie_file

    try:
        flat_opts = {**base_opts, "extract_flat": "in_playlist", "skip_download": True}
        with yt_dlp.YoutubeDL(flat_opts) as ydl:
            info = ydl.extract_info(target, download=False)
            candidates = (
                [e for e in info["entries"] if isinstance(e, dict)]
                if isinstance(info, dict) and "entries" in info and info["entries"]
                else ([info] if isinstance(info, dict) else [])
            )

            if not candidates:
                raise RuntimeError("YouTube qidiruvidan natija qaytmadi.")

            safe_candidates = []
            for cand in candidates:
                c_title = cand.get("title") or ""
                c_channel = cand.get("uploader") or cand.get("channel") or ""
                c_desc = cand.get("description") or ""
                if is_prohibited_uzb_content(c_title, c_channel, c_desc):
                    continue

                vid_id = cand.get("id") or ""
                raw_url = cand.get("url") or cand.get("webpage_url") or ""
                if raw_url.startswith("http"):
                    v_url = raw_url
                elif vid_id:
                    v_url = f"https://www.youtube.com/watch?v={vid_id}"
                else:
                    v_url = query_or_url if is_direct_url else ""

                safe_candidates.append(
                    {
                        "raw": cand,
                        "title": c_title or "YouTube Media",
                        "channel": c_channel,
                        "url": v_url,
                    }
                )

            if not safe_candidates:
                raise PermissionError(
                    "Topilgan kanal yoki video O‘zbekiston Respublikasida taqiqlangan manbalar ro‘yxatiga kiradi."
                )

            variants = safe_candidates[:5]
            selected_item = variants[0]
            title = selected_item["title"]
            video_url = selected_item["url"]

        if link_only:
            return None, title, video_url, variants

        if audio_only:
            dl_opts = {
                **base_opts,
                "format": "bestaudio/best",
                "outtmpl": out_tmpl,
            }
        else:
            dl_opts = {
                **base_opts,
                "format": "best[ext=mp4]/best[height<=720]/best",
                "outtmpl": out_tmpl,
            }

        try:
            with yt_dlp.YoutubeDL(dl_opts) as ydl:
                dl_info = ydl.extract_info(video_url or target, download=True)
                if isinstance(dl_info, dict) and "entries" in dl_info and dl_info["entries"]:
                    dl_info = dl_info["entries"][0]

                full_title = dl_info.get("title") or title
                full_channel = dl_info.get("uploader") or dl_info.get("channel") or ""
                full_desc = dl_info.get("description") or ""
                if is_prohibited_uzb_content(full_title, full_channel, full_desc):
                    raise PermissionError(
                        "Ushbu ma’ruza, kanal yoki video O‘zbekiston Respublikasida taqiqlangan, yuklash bekor qilindi."
                    )

                title = full_title
                downloaded_path = ydl.prepare_filename(dl_info)
                if not os.path.exists(downloaded_path):
                    files = [
                        p
                        for p in glob.glob(os.path.join(out_dir, "*"))
                        if os.path.abspath(p) != os.path.abspath(cookie_file or "")
                    ]
                    if not files:
                        return None, title, video_url, variants
                    downloaded_path = files[0]

                return downloaded_path, title, video_url, variants
        except PermissionError:
            raise
        except Exception as dl_err:
            logging.warning(f"Media fayl yuklanmadi, variantlar va havola qaytariladi: {dl_err}")
            return None, title, video_url, variants
    finally:
        if cookie_file and os.path.exists(cookie_file):
            try:
                os.remove(cookie_file)
            except OSError:
                pass


# --- 6. Eslatmalar, YouTube Akkaunt va Fayl Yaratish (Gemini Function Calling Tools) ---
def _safe_filename(filename: str, default_ext: str) -> str:
    clean = re.sub(r'[\\/*?:"<>|]', "_", filename.strip())
    if not clean:
        clean = f"jarvis_hujjat{default_ext}"
    if not clean.lower().endswith(default_ext):
        clean += default_ext
    return clean


def create_tools_for_user(
    user_id: int,
    generated_files: list,
    yt_download_requests: list,
):
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

    def get_my_youtube_videos(section: str = "watch_later", limit: int = 10) -> str:
        """Foydalanuvchining shaxsiy YouTube akkauntidan videolar ro'yxatini oladi."""
        try:
            feed_label, items = fetch_my_youtube_entries(section=section, limit=limit)
            if not items:
                return f"YouTube akkauntingizning '{feed_label}' bo‘limida ruxsat etilgan videolar topilmadi."

            lines = [f"YouTube akkauntingiz — {feed_label} (oxirgi {len(items)} ta):"]
            for item in items:
                channel_part = f" ({item['uploader']})" if item["uploader"] else ""
                lines.append(
                    f"{item['index']}. {item['title']}{channel_part}\n   Havola: {item['url']}"
                )
            return "\n".join(lines)
        except Exception as e:
            return f"YouTube akkaunt ma'lumotlarini olishda xatolik: {e}"

    def download_from_my_youtube(
        section: str = "watch_later",
        index: int = 1,
        audio_only: bool = False,
    ) -> str:
        """Foydalanuvchining shaxsiy YouTube akkauntidagi ko'rsatilgan bo'lim va tartib raqamidagi videoni yuklab yuborish navbatiga qo'shadi."""
        try:
            idx = max(1, int(index))
            feed_label, items = fetch_my_youtube_entries(section=section, limit=idx)
            if len(items) < idx:
                return f"Xatolik: '{feed_label}' bo‘limida {idx}-raqamli video topilmadi."

            chosen = items[idx - 1]
            if not chosen["url"]:
                return f"Xatolik: '{chosen['title']}' videosining havolasi topilmadi."

            yt_download_requests.append(
                {
                    "url": chosen["url"],
                    "title": chosen["title"],
                    "audio_only": bool(audio_only),
                }
            )
            media_type = "audio" if audio_only else "video"
            return (
                f"'{feed_label}' ro‘yxatidagi #{idx} ('{chosen['title']}') {media_type} sifatida "
                f"yuklash navbatiga qo‘shildi. Havola: {chosen['url']}"
            )
        except Exception as e:
            return f"Akkauntdan video tanlashda xatolik: {e}"

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
        get_my_youtube_videos,
        download_from_my_youtube,
        create_excel_file,
        create_word_file,
        create_pdf_file,
        create_text_file,
    ]


def should_enable_custom_tools(text: str) -> bool:
    """So'rov eslatma, fayl yaratish yoki shaxsiy YouTube akkauntga oid ekanligini tekshiradi."""
    lower_text = text.lower()
    return (
        any(k in lower_text for k in REMINDER_KEYWORDS)
        or any(k in lower_text for k in FILE_GEN_KEYWORDS)
        or any(k in lower_text for k in MY_YT_KEYWORDS)
    )


def build_chat_config(user_id: int, custom_tools: list | None = None):
    now_str = datetime.now(LOCAL_TZ).strftime("%Y-%m-%d %H:%M:%S (%A)")
    dynamic_instruction = (
        f"{JARVIS_BASE_INSTRUCTION}\n"
        f"Hozirgi aniq mahalliy sana va vaqt ({TIMEZONE_STR}): {now_str}. "
        f"Bugungi kun va soat holatidagi savollarga shu vaqt hamda Google Search asosida javob ber."
    )

    active_tools = (
        custom_tools
        if custom_tools
        else [types.Tool(google_search=types.GoogleSearch())]
    )

    return types.GenerateContentConfig(
        system_instruction=dynamic_instruction,
        tools=active_tools,
        temperature=0.7,
    )


async def send_message_with_fallback(
    user_id: int, message_content: Any, use_custom_tools: bool = False
):
    """Zaxira modellar ro'yxati bo'yicha jonli qidiruv yoki maxsus funksiyalar bilan javob oladi."""
    history = load_history_from_db(user_id)
    last_error = None
    quota_hit = False

    for model_id in FALLBACK_MODELS:
        generated_files: list[tuple[str, bytes]] = []
        yt_download_requests: list[dict] = []
        custom_tools = (
            create_tools_for_user(user_id, generated_files, yt_download_requests)
            if use_custom_tools
            else None
        )
        config = build_chat_config(user_id, custom_tools=custom_tools)

        try:
            chat = client.chats.create(
                model=model_id,
                history=history,
                config=config,
            )
            response = await asyncio.to_thread(chat.send_message, message_content)
            return response, generated_files, yt_download_requests
        except Exception as e:
            last_error = e
            err_str = str(e)
            if "429" in err_str or "RESOURCE_EXHAUSTED" in err_str:
                quota_hit = True
            logging.warning(
                f"{model_id} vaqtincha javob bermadi ({e}), keyingi zaxira modelga o'tilmoqda..."
            )
            await asyncio.sleep(1.2)
            continue

    if quota_hit:
        raise RuntimeError(
            "Gemini API bepul kvotasi vaqtincha tugadi (429). Birozdan so‘ng qayta urinib ko‘ring "
            "yoki Google AI Studio’dan yangi API kalit ulab qo‘ying."
        )
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


# --- 7. Orqa fon Eslatma Tekshiruvchisi ---
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


# --- 8. Matn, Ovoz va Fayl Yuborish Yordamchilari ---
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


def _format_variants_text(variants: list[dict]) -> str:
    if not variants:
        return ""
    lines = ["\n\nTopilgan varyantlar:"]
    for idx, item in enumerate(variants[:5], start=1):
        ch_part = f" — {item['channel']}" if item.get("channel") else ""
        url_part = f"\n   🔗 {item['url']}" if item.get("url") else ""
        lines.append(f"{idx}. {item['title']}{ch_part}{url_part}")
    return "\n".join(lines)


async def process_yt_download_and_send(
    message: Message, yt_target: str, audio_only: bool, link_only: bool = False
) -> str:
    """YouTube havolasi yoki qidiruv so'zini tekshirib, imkon qadar faylni va variantlar havolasini yuboradi."""
    status_msg = await message.answer(
        f"Janob, '{yt_target}' bo‘yicha YouTube variantlari qidirilmoqda..."
        if link_only
        else (
            f"Janob, '{yt_target}' bo‘yicha audio va variantlar tayyorlanmoqda, biroz kuting..."
            if audio_only
            else f"Janob, '{yt_target}' bo‘yicha video va variantlar tayyorlanmoqda, biroz kuting..."
        )
    )
    await bot.send_chat_action(
        chat_id=message.chat.id,
        action="typing" if link_only else ("upload_voice" if audio_only else "upload_video"),
    )

    try:
        with tempfile.TemporaryDirectory() as tmp_dir:
            file_path, title, video_url, variants = await asyncio.to_thread(
                download_youtube_media, yt_target, audio_only, tmp_dir, link_only
            )
            variants_text = _format_variants_text(variants)

            if link_only or not file_path:
                main_link = f"\n🔗 Havola: {video_url}" if video_url else ""
                await status_msg.edit_text(f"🎬 {title}{main_link}{variants_text}")
                return f"YouTube’dan '{title}' bo‘yicha variantlar yuborildi: {video_url}"

            file_size_mb = os.path.getsize(file_path) / (1024 * 1024)
            url_part = f"\n🔗 Havola: {video_url}" if video_url else ""

            if file_size_mb > 49.5:
                await status_msg.edit_text(
                    f"Janob, '{title}' fayli hajmi ({file_size_mb:.1f} MB) Telegram botining 50 MB cheklovidan katta, "
                    f"shuning uchun faylni jo‘nata olmadim. Quyidagi havola va varyantlardan foydalanishingiz mumkin:{url_part}{variants_text}"
                )
                return f"YouTube’dan '{title}' havolasi va variantlari berildi: {video_url}"

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

            if len(variants) > 1:
                await status_msg.edit_text(f"Boshqa varyantlar:{variants_text}")
            else:
                await status_msg.delete()

            return f"YouTube’dan '{title}' ({'audio' if audio_only else 'video'}) va variantlari yuborildi: {video_url}"
    except Exception as e:
        await status_msg.edit_text(f"Uzr so‘rayman, Janob. {e}")
        raise


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


# --- 9. Telegram Handlerlar ---
@dp.message(CommandStart())
async def cmd_start(message: Message):
    user_id = message.from_user.id
    clear_history_in_db(user_id)

    welcome_text = (
        "Xush kelibsiz, Janob. J.A.R.V.I.S. tizimi to‘liq ishga tushirildi: "
        "bugungi kun va soat holatidagi jonli ma’lumotlar, YouTube hamda ochiq Telegram kanallaridan "
        "qidirish va yuklash (O‘zbekiston Respublikasi qonunchiligida taqiqlangan ma’ruza va kanallar filtrlangan holda), "
        "ovozli aloqa, eslatmalar hamda Excel, Word, PDF va TXT hujjatlarni tahlil qilish va yaratish tayyor."
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
            response, generated_files, _ = await send_message_with_fallback(
                user_id, full_prompt, use_custom_tools=True
            )
        else:
            doc_part = types.Part.from_bytes(
                data=doc_bytes,
                mime_type=gemini_mime,
            )
            response, generated_files, _ = await send_message_with_fallback(
                user_id, [doc_part, prompt_text], use_custom_tools=True
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

        response, generated_files, yt_requests = await send_message_with_fallback(
            user_id,
            [
                audio_part,
                "Foydalanuvchining ovozli xabarini tingla va unga J.A.R.V.I.S. sifatida javob ber. Agar fayl yaratish yoki YouTube akkauntdan video olish/yuklash so'ralgan bo'lsa, tegishli funksiyadan foydalan.",
            ],
            use_custom_tools=True,
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
        for req in yt_requests:
            await process_yt_download_and_send(
                message, req["url"], req["audio_only"], link_only=False
            )

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

        response, generated_files, _ = await send_message_with_fallback(
            user_id, [image_part, prompt_text], use_custom_tools=True
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
    lower_text = text_body.lower()

    # 0. O'zbekiston Respublikasida taqiqlangan ma'ruzalar, ekstremistik oqimlar va kanallarni bloklash
    if is_prohibited_uzb_content(text_body):
        await message.answer(
            "Uzr so‘rayman, Janob. O‘zbekiston Respublikasi qonunchiligida taqiqlangan "
            "ma’ruzalar, ekstremistik oqimlar yoki bloklangan kanal va videolarni qidira olmayman hamda yuklab bera olmayman."
        )
        return

    # 1. Ochiq Telegram kanali (t.me/... yoki @username) yoki Telegram qidiruv so'rovini tekshirish
    tg_url_match = TELEGRAM_URL_PATTERN.search(text_body)
    at_channel_match = re.search(r"@([a-zA-Z0-9_]{4,})", text_body)
    is_telegram_query = bool(tg_url_match) or (
        any(k in lower_text for k in TELEGRAM_SEARCH_KEYWORDS)
    )

    if is_telegram_query:
        await bot.send_chat_action(chat_id=message.chat.id, action="typing")
        target_channel = (
            tg_url_match.group(1)
            if tg_url_match
            else (at_channel_match.group(1) if at_channel_match else "")
        )
        clean_tg_kw = re.sub(
            r"(https?://\S+|@[a-zA-Z0-9_]+|\b(jarvis|telegram|telegramdan|telegramdagi|kanaldan|kanallardan|dan|dagi|oxirgi|yangi|postlarni|xabarlarni|qo'shiq|qoshiq|qo'shiqlarini|qoshiqlarini|video|videolarini|topib|top|ber|o'qib|oqib|chiqar|menga)\b)",
            " ",
            lower_text,
        )
        clean_tg_kw = re.sub(r"\s+", " ", clean_tg_kw).strip()

        if target_channel:
            try:
                posts = await fetch_public_telegram_channel_posts(
                    target_channel, keyword=clean_tg_kw, limit=6
                )
                if posts:
                    lines = [f"📢 Ochiq Telegram kanali ({posts[0]['channel']}) natijalari:"]
                    for idx, p in enumerate(posts, start=1):
                        lines.append(f"\n{idx}. {p['text']}\n   🔗 {p['url']}")
                    reply_text = "\n".join(lines)
                    save_message_to_db(user_id, "user", text_body)
                    save_message_to_db(user_id, "model", reply_text)
                    await send_jarvis_reply(message, reply_text)
                    return
            except PermissionError as pe:
                await message.answer(f"Uzr so‘rayman, Janob. {pe}")
                return
            except Exception as tg_err:
                logging.warning(f"Ochiq Telegram kanalini o'qishda xatolik: {tg_err}")

        tg_search_prompt = (
            f"Ochiq Telegram kanallari va postlari (site:t.me) ichidan quyidagi so'rov bo'yicha "
            f"eng mos kanallar/postlar va ularning haqiqiy t.me havolalarini ro'yxat qilib ber: {clean_tg_kw or text_body}"
        )
        try:
            response, _, _ = await send_message_with_fallback(
                user_id, tg_search_prompt, use_custom_tools=False
            )
            reply_text = (
                format_response_with_sources(response)
                or "Janob, ochiq Telegram kanallaridan mos natija topilmadi."
            )
            save_message_to_db(user_id, "user", text_body)
            save_message_to_db(user_id, "model", reply_text)
            await send_jarvis_reply(message, reply_text)
        except Exception as e:
            await message.answer(f"Uzr so‘rayman, Janob. {e}")
        return

    # 2. Shaxsiy YouTube akkaunt so'rovlarini aniqlash
    is_my_yt_query = any(k in lower_text for k in MY_YT_KEYWORDS) and any(
        w in lower_text
        for w in ("youtube", "yutub", "yutube", "video", "ro'yxat", "royxat", "yukl", "tort")
    )

    # 3. Har qanday YouTube media (qo'shiq, video, mp3, mp4, audio, silka, varyantlar, yuklab/tortib ber) so'rovi
    yt_match = YOUTUBE_URL_PATTERN.search(text_body)
    has_media_download_intent = (
        not is_my_yt_query
        and not any(k in lower_text for k in FILE_GEN_KEYWORDS)
        and any(w in lower_text for w in YT_ACTION_KEYWORDS)
    )

    if yt_match or has_media_download_intent:
        if yt_match:
            yt_target = yt_match.group(1)
        else:
            clean_query = re.sub(
                r"\b(jarvis|youtube|yutub|yutube|dan|ni|bironta|bir|video|videosini|videoni|qo'shiq|qoshiq|qo'shig'ini|qoshig'ini|qoʻshigʻini|to'plam|toplam|to'plamini|toplamini|musiqa|audio|mp3|mp4|topib|top|yuklab|yuklaver|yukla|tortib|tort|skachat|tashlab|tashla|ber|bersang|bersinda|qilib|shuni|silkasini|silka|havolasini|havola|linkini|link|varyant|variant|varyantlar|variantlar|ha|ham|menga|desam|kerak|bo'laveradi|bolaveradi)\b",
                " ",
                lower_text,
            )
            clean_query = re.sub(r"\s+", " ", clean_query).strip()
            yt_target = clean_query or text_body

        audio_only = any(k in lower_text for k in AUDIO_INTENT_KEYWORDS)
        link_only = any(k in lower_text for k in LINK_ONLY_KEYWORDS) and not any(
            k in lower_text for k in ("yukl", "tort", "skachat", "mp3", "mp4", "audio", "video", "fayl", "to'plam", "toplam")
        )

        try:
            log_summary = await process_yt_download_and_send(
                message, yt_target, audio_only, link_only
            )
            save_message_to_db(user_id, "user", text_body)
            save_message_to_db(user_id, "model", log_summary)
        except Exception as e:
            logging.exception("YouTube yuklashda xatolik:")
        return

    # 4. Bugungi kun va soat holatidagi jonli ma'lumotlar (Google Search), eslatmalar, hujjatlar
    await bot.send_chat_action(chat_id=message.chat.id, action="typing")

    try:
        needs_custom_tools = should_enable_custom_tools(text_body)
        response, generated_files, yt_requests = await send_message_with_fallback(
            user_id, text_body, use_custom_tools=needs_custom_tools
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
        for req in yt_requests:
            await process_yt_download_and_send(
                message, req["url"], req["audio_only"], link_only=False
            )

    except Exception as e:
        logging.exception("Matn so'rovida xatolik:")
        await message.answer(
            f"Uzr so‘rayman, Janob. {e}"
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