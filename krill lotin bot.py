# krill_lotin_bot_full.py
import logging
import os
import re
import tempfile
import shutil
from telegram import Update, ReplyKeyboardMarkup
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    filters,
    ContextTypes,
)
from docx import Document

# --- CONFIG ---
TOKEN = "8451093366:AAFr9rkhFJYV060TPcjB069iKDd6sObezXQ"

# --- Logging ---
logging.basicConfig(format="%(asctime)s - %(levelname)s - %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

# --- Helper functions ---
def normalize_apostrophes(text: str) -> str:
    if not text:
        return text
    for a in ["'", "‘", "’", "ʼ", "ʻ", "`", "ʹ", "´"]:
        text = text.replace(a, "'")
    return text

def apply_apostrophe_rules(s: str) -> str:
    s = re.sub(r"o'", "O_TEMP", s, flags=re.IGNORECASE)
    s = re.sub(r"g'", "G_TEMP", s, flags=re.IGNORECASE)
    s = s.replace("'", "ъ")
    s = s.replace("O_TEMP", "ў")
    s = s.replace("G_TEMP", "ғ")
    return s

# --- Transliteration maps ---
LATIN_TO_CYR = {
    "sh": "ш", "ch": "ч", "ts": "ц",
    "a":"а","b":"б","d":"д","e":"е","f":"ф","g":"г","h":"ҳ","i":"и",
    "j":"ж","k":"к","l":"л","m":"м","n":"н","o":"о","p":"п","q":"қ","r":"р",
    "s":"с","t":"т","u":"у","v":"в","x":"х","y":"й","z":"з",
    "yo": "ё", "ya": "я", "yu": "ю", "ye": "е",
}
CYR_TO_LAT = {v:k for k,v in LATIN_TO_CYR.items()}

def expand_case(d):
    out = {}
    for k,v in d.items():
        out[k] = v
        out[k.upper()] = v.upper()
        if len(k) > 1:
            out[k.capitalize()] = v.capitalize()
    return out

LATIN_TO_CYR = expand_case(LATIN_TO_CYR)
CYR_TO_LAT = expand_case(CYR_TO_LAT)

LAT_TO_CYR_RE = re.compile("|".join(sorted(map(re.escape, LATIN_TO_CYR.keys()), key=len, reverse=True)))
CYR_TO_LAT_RE = re.compile("|".join(sorted(map(re.escape, CYR_TO_LAT.keys()), key=len, reverse=True)))

def replace_match_case(m, mapping):
    t = m.group(0)
    return mapping.get(t, t)

# --- Force e rule ---
def force_e_rule(text: str) -> str:
    """Har bir so'z boshidagi e -> э"""
    def repl(match):
        word = match.group(0)
        if word[0].lower() == 'e':
            return ('Э' if word[0].isupper() else 'э') + word[1:]
        return word
    return re.sub(r'\b\w+\b', repl, text)

# --- Transliteration ---
def transliterate_text(s: str, to_cyr=True) -> str:
    if not s:
        return s
    s = normalize_apostrophes(s)
    s = apply_apostrophe_rules(s)

    if to_cyr:
        # 1️⃣ So'z boshidagi e -> э
        s = force_e_rule(s)
        # 2️⃣ Keyin boshqa transliteratsiya
        tmp = LAT_TO_CYR_RE.sub(lambda m: replace_match_case(m, LATIN_TO_CYR), s)
        return tmp
    else:
        tmp = CYR_TO_LAT_RE.sub(lambda m: replace_match_case(m, CYR_TO_LAT), s)
        return tmp

# --- DOCX conversion ---
def convert_docx_preserve_format(input_path_or_stream, output_path_or_stream, to_cyr=True):
    doc = Document(input_path_or_stream)
    for para in doc.paragraphs:
        for run in para.runs:
            if run.text:
                run.text = transliterate_text(run.text, to_cyr=to_cyr)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        if run.text:
                            run.text = transliterate_text(run.text, to_cyr=to_cyr)
    try:
        for section in doc.sections:
            if section.header:
                for para in section.header.paragraphs:
                    for run in para.runs:
                        if run.text:
                            run.text = transliterate_text(run.text, to_cyr=to_cyr)
            if section.footer:
                for para in section.footer.paragraphs:
                    for run in para.runs:
                        if run.text:
                            run.text = transliterate_text(run.text, to_cyr=to_cyr)
    except Exception:
        logger.debug("Header/footer skipped.")
    if isinstance(output_path_or_stream, str):
        doc.save(output_path_or_stream)
        return output_path_or_stream
    else:
        return doc

# --- Telegram handlers ---
START_TEXT = "Assalomu alaykum!\nMen Lotin ↔ Krill botiman.\nFayl yuboring (.docx) yoki matn yozing."

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    kb = [["Lotindan Krillga", "Krilldan Lotinga"]]
    await update.message.reply_text(START_TEXT, reply_markup=ReplyKeyboardMarkup(kb, resize_keyboard=True))

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    mode = context.user_data.get("mode")
    if not mode:
        await update.message.reply_text("Avval yo‘nalishni tanlang: Lotindan Krillga yoki Krilldan Lotinga.")
        return
    doc = update.message.document
    if not doc:
        await update.message.reply_text("Fayl topilmadi. Iltimos .docx fayl yuboring.")
        return
    fname = (doc.file_name or "").strip()
    fname_lower = fname.lower()
    mime = getattr(doc, "mime_type", "") or ""
    if fname_lower.endswith(".doc") or "msword" in mime:
        await update.message.reply_text("⚠️ Men .doc fayllarni o‘qiy olmayman. Iltimos .docx yuboring.")
        return
    if not (fname_lower.endswith(".docx") or "wordprocessingml" in mime):
        await update.message.reply_text("❗ Faqat .docx fayllarni qabul qilaman.")
        return
    file = await context.bot.get_file(doc.file_id)
    tmp_dir = None
    try:
        tmp_dir = tempfile.mkdtemp(prefix="tg_doc_")
        in_path = os.path.join(tmp_dir, fname)
        out_path = os.path.join(tmp_dir, fname)
        await file.download_to_drive(in_path)
        try:
            convert_docx_preserve_format(in_path, out_path, to_cyr=(mode=="to_cyr"))
        except Exception:
            await update.message.reply_text("❌ Faylni o'zgartirishda xatolik yuz berdi.")
            return
        with open(out_path, "rb") as f:
            await update.message.reply_document(f, filename=fname)
    except Exception as e:
        logger.exception("handle_document error:")
        await update.message.reply_text("Kutilmagan xatolik yuz berdi.")
    finally:
        if tmp_dir:
            shutil.rmtree(tmp_dir, ignore_errors=True)

async def text_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text or ""
    if text == "Lotindan Krillga":
        context.user_data["mode"] = "to_cyr"
        await update.message.reply_text("✅ Lotindan Krillga rejimi tanlandi.")
        return
    if text == "Krilldan Lotinga":
        context.user_data["mode"] = "to_lat"
        await update.message.reply_text("✅ Krilldan Lotinga rejimi tanlandi.")
        return
    mode = context.user_data.get("mode")
    if not mode:
        await update.message.reply_text("Avval yo‘nalishni tanlang.")
        return
    to_cyr = (mode=="to_cyr")
    result = transliterate_text(text, to_cyr=to_cyr)
    await update.message.reply_text(result)

async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    logger.exception("Unhandled exception:")
    try:
        if hasattr(update, "message") and update.message:
            await update.message.reply_text("Kutilmagan xatolik yuz berdi.")
    except Exception:
        pass

def main():
    app = ApplicationBuilder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, text_message))
    app.add_error_handler(error_handler)
    logger.info("Bot ishga tushdi...")
    app.run_polling()

if __name__ == "__main__":
    main()
