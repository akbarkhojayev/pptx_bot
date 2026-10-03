import asyncio
import logging
import os
import shutil
import tempfile

from aiogram import Bot, Dispatcher, F, types
from aiogram.enums import ChatAction
from aiogram.filters import Command, CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import FSInputFile, KeyboardButton, ReplyKeyboardMarkup, ReplyKeyboardRemove

from ai_content import generate_deck
from pptx_builder import create_presentation_file, normalize_image

BOT_TOKEN = os.environ.get("BOT_TOKEN", "")
if not BOT_TOKEN:
    raise SystemExit("BOT_TOKEN environment o'zgaruvchisi o'rnatilmagan (.env.example ga qarang).")

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()
logging.basicConfig(level=logging.INFO)

MAX_TOPIC_LEN = 150
MAX_IMAGES = 8

BTN_SKIP = "⏭ Rasmsiz yaratish"
BTN_LARGE = "🖼 Alohida katta slayd"
BTN_TEXT = "📝 Matn yonida"
BTN_DONE = "✅ Tayyor, yaratish"
BTN_MORE = "➕ Yana rasm"


def _kb(*rows) -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text=t) for t in row] for row in rows],
        resize_keyboard=True,
        one_time_keyboard=True,
    )


class Form(StatesGroup):
    waiting_for_topic = State()
    waiting_for_image = State()
    waiting_for_type = State()
    waiting_for_more = State()


def _norm(text) -> str:
    return (text or "").lower().strip()


async def _user_tmp_dir(state: FSMContext) -> str:
    data = await state.get_data()
    path = data.get("tmp_dir")
    if not path or not os.path.isdir(path):
        path = tempfile.mkdtemp(prefix="pptx_bot_")
        await state.update_data(tmp_dir=path)
    return path


def _cleanup_dir(path) -> None:
    if path:
        shutil.rmtree(path, ignore_errors=True)


# ---------------- BOT HANDLERS ----------------
@dp.message(CommandStart())
async def cmd_start(message: types.Message, state: FSMContext):
    _cleanup_dir((await state.get_data()).get("tmp_dir"))
    await state.clear()
    await message.answer(
        "👋 Salom! Men Prezentatsiya Botman.\n\n"
        "Mavzu yuboring — men reja, slaydlar (punktlar, raqamlar, xronologiya, taqqoslash) va har bir slayd "
        "uchun ma'ruzachi eslatmalari bilan tayyor PowerPoint fayl yasab beraman.\n\n"
        "Prezentatsiya mavzusini kiriting:",
        reply_markup=ReplyKeyboardRemove(),
    )
    await state.set_state(Form.waiting_for_topic)


@dp.message(Command("cancel"))
async def cmd_cancel(message: types.Message, state: FSMContext):
    _cleanup_dir((await state.get_data()).get("tmp_dir"))
    await state.clear()
    await message.answer("Bekor qilindi. Qaytadan boshlash uchun /start", reply_markup=ReplyKeyboardRemove())


@dp.message(Form.waiting_for_topic)
async def handle_topic_input(message: types.Message, state: FSMContext):
    topic = (message.text or "").strip()
    if not topic:
        await message.reply("Iltimos, mavzuni matn ko'rinishida yozing.")
        return
    if len(topic) > MAX_TOPIC_LEN:
        await message.reply(f"Mavzu juda uzun ({len(topic)} belgi). Iltimos, {MAX_TOPIC_LEN} belgigacha qisqartiring.")
        return
    await state.update_data(topic=topic, slide_images=[], large_images=[])
    await message.answer(
        "✅ Mavzu saqlandi.\n\n"
        "Endi taqdimotga qo'shiladigan rasm yuboring yoki rasmsiz davom eting:",
        reply_markup=_kb([BTN_SKIP]),
    )
    await state.set_state(Form.waiting_for_image)


async def _receive_image(message: types.Message, state: FSMContext):
    data = await state.get_data()
    count = len(data.get("slide_images", [])) + len(data.get("large_images", []))
    if count >= MAX_IMAGES:
        await message.reply(f"Ko'pi bilan {MAX_IMAGES} ta rasm qo'shish mumkin.", reply_markup=_kb([BTN_DONE]))
        await state.set_state(Form.waiting_for_more)
        return
    if message.photo:
        file_id = message.photo[-1].file_id
    else:
        file_id = message.document.file_id
    tmp_dir = await _user_tmp_dir(state)
    raw_path = os.path.join(tmp_dir, f"raw_{count}")
    img_path = os.path.join(tmp_dir, f"image_{count}.jpg")
    try:
        file = await bot.get_file(file_id)
        await bot.download_file(file.file_path, raw_path)
        await asyncio.get_running_loop().run_in_executor(None, normalize_image, raw_path, img_path)
    except Exception:
        logging.exception("Rasmni yuklab olishda xatolik")
        await message.reply("❌ Rasmni o'qib bo'lmadi. Boshqa rasm yuboring yoki davom eting.")
        return
    finally:
        if os.path.exists(raw_path):
            os.remove(raw_path)
    await state.update_data(last_img_path=img_path)
    await message.reply(
        "🖼 Rasm yuklandi! Uni qayerga joylay?\n"
        f"• {BTN_LARGE} — rasm butun slaydni egallaydi\n"
        f"• {BTN_TEXT} — rasm matnli slaydning o'ng tomonida",
        reply_markup=_kb([BTN_LARGE, BTN_TEXT]),
    )
    await state.set_state(Form.waiting_for_type)


_image_filter = F.photo | (F.document & F.document.mime_type.startswith("image/"))


@dp.message(_image_filter, Form.waiting_for_image)
async def handle_photo(message: types.Message, state: FSMContext):
    await _receive_image(message, state)


@dp.message(_image_filter, Form.waiting_for_more)
async def handle_additional_photo(message: types.Message, state: FSMContext):
    await _receive_image(message, state)


@dp.message(Form.waiting_for_image)
async def handle_skip_image(message: types.Message, state: FSMContext):
    if _norm(message.text) in ("skip", _norm(BTN_SKIP)):
        await _finish(message, state)
    else:
        await message.reply("Iltimos, rasm yuboring yoki rasmsiz davom eting.", reply_markup=_kb([BTN_SKIP]))


@dp.message(Form.waiting_for_type)
async def handle_type(message: types.Message, state: FSMContext):
    data = await state.get_data()
    img_path = data.get("last_img_path")
    if not img_path:
        await message.reply("Xatolik yuz berdi. Qaytadan boshlang: /start", reply_markup=ReplyKeyboardRemove())
        await state.clear()
        return
    typ = _norm(message.text)
    if typ in ("asosiy", _norm(BTN_LARGE)):
        key = "large_images"
    elif typ in ("matn", _norm(BTN_TEXT)):
        key = "slide_images"
    else:
        await message.reply("Iltimos, quyidagi tugmalardan birini tanlang.", reply_markup=_kb([BTN_LARGE, BTN_TEXT]))
        return
    await state.update_data(**{key: data.get(key, []) + [img_path]}, last_img_path=None)
    await message.reply("✅ Rasm saqlandi! Yana rasm yuboring yoki taqdimotni yaratishni boshlang.",
                        reply_markup=_kb([BTN_DONE]))
    await state.set_state(Form.waiting_for_more)


@dp.message(Form.waiting_for_more)
async def handle_more(message: types.Message, state: FSMContext):
    if _norm(message.text) in ("done", "tayyor", _norm(BTN_DONE)):
        await _finish(message, state)
    else:
        await message.reply("Yana rasm yuboring yoki tugmani bosing.", reply_markup=_kb([BTN_DONE]))


@dp.message()
async def handle_unknown(message: types.Message):
    await message.reply("Boshlash uchun /start ni bosing.")


async def _finish(message: types.Message, state: FSMContext):
    data = await state.get_data()
    await state.clear()
    tmp_dir = data.get("tmp_dir") or tempfile.mkdtemp(prefix="pptx_bot_")
    try:
        await create_presentation(message, data["topic"], data.get("slide_images", []),
                                  data.get("large_images", []), tmp_dir)
    finally:
        _cleanup_dir(tmp_dir)


async def create_presentation(message: types.Message, topic: str, slide_images: list, large_images: list,
                              tmp_dir: str):
    loop = asyncio.get_running_loop()
    await message.answer(f"🔍 «{topic}» bo'yicha ma'lumot to'planmoqda va reja tuzilmoqda...",
                         reply_markup=ReplyKeyboardRemove())
    await bot.send_chat_action(message.chat.id, ChatAction.TYPING)
    deck, source = await loop.run_in_executor(None, generate_deck, topic)
    if not deck:
        await message.answer("😔 Bu mavzu bo'yicha ma'lumot topilmadi. Mavzuni aniqroq yozib qaytadan urining: /start")
        return

    await message.answer("🎨 Slaydlar dizayni tayyorlanmoqda...")
    await bot.send_chat_action(message.chat.id, ChatAction.UPLOAD_DOCUMENT)
    try:
        pptx_path, total = await loop.run_in_executor(None, create_presentation_file, deck, tmp_dir, slide_images,
                                               large_images)
    except Exception:
        logging.exception("PPTX yaratishda xatolik")
        await message.answer("❌ Faylni yaratishda xatolik yuz berdi. Qaytadan urinib ko'ring: /start")
        return

    note = "" if source == "ai" else "\nℹ️ AI vaqtincha ishlamadi, kontent Wikipedia'dan olindi."
    await message.answer_document(
        FSInputFile(pptx_path),
        caption=f"✅ «{deck['title']}» — {total} ta slayd.\n"
                f"💬 Har bir slaydning ma'ruzachi eslatmalari (Notes) bor.{note}\n\nYangi taqdimot: /start",
    )


if __name__ == "__main__":
    logging.info("🚀 Bot ishga tushmoqda...")
    asyncio.run(dp.start_polling(bot))
