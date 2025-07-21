import asyncio
import os
from aiogram import Bot, Dispatcher, types, F
from aiogram.enums import ParseMode
from aiogram.client.default import DefaultBotProperties
from aiogram.fsm.context import FSMContext
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton
from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton, CallbackQuery
from aiogram.types import Message, FSInputFile
from aiogram.filters import Command
from aiogram.fsm.state import StatesGroup, State
from aiogram.fsm.storage.memory import MemoryStorage
import wikipedia
import logging

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler("bot.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)
wikipedia.set_lang('uz')
API_TOKEN = "7890416134:AAFMbGQieQoJiGO-3HkKC8GA0T0j6RGor-o"
ADMIN_ID = "5922081119"

bot = Bot(token=API_TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
dp = Dispatcher(storage=MemoryStorage())

main_menu = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="About"), KeyboardButton(text="Projects")],
        [KeyboardButton(text="Contact"), KeyboardButton(text="Skils")],
        [KeyboardButton(text="📄 CV"), KeyboardButton(text="✉️ Fikr bildirish")],
    ],
    resize_keyboard=True
)
cancel_keyboard = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="❌ Bekor qilish")]
    ],
    resize_keyboard=True
)

contact_keyboard = InlineKeyboardMarkup(
    inline_keyboard=[
        [InlineKeyboardButton(text="Telegram", url="https://t.me/maxmudov_1117")],
        # [InlineKeyboardButton(text="Gmail", url="bekmurodmaxmudov122@gmail.com")],
        [InlineKeyboardButton(text="LinkedIn", url="https://linkedin.com/in/BekmurodMaxmudov")],
        [InlineKeyboardButton(text="GitHub", url="https://github.com/maxmudov1117")]
    ]
)

projects_keyboard = InlineKeyboardMarkup(
    inline_keyboard=[
        [
            InlineKeyboardButton(text="KinoBot", callback_data='KinoBot'),
            InlineKeyboardButton(text="Yordam+", callback_data='Yordam+'),
            InlineKeyboardButton(text="TarjimonBot", callback_data='TarjimonBot')
        ]
    ]
)


@dp.message(Command("start"))
async def cmd_start(message: Message):
    logger.info(f"Foydalanuvchi: {message.from_user.full_name} ({message.from_user.id}) /start yubordi")
    await message.answer("""
    👋 Assalomu alaykum!  
Men Bekmurod Maxmudovning rasmiy portfolio botiman.  
Quyidagi buyruqlar orqali men haqimdagi ma’lumotlarni olishingiz mumkin:
    """, reply_markup=main_menu)


class CVForm(StatesGroup):
    waiting_for_name = State()
    waiting_for_surname = State()
    waiting_for_phone = State()


@dp.message(F.text == "📄 CV")
async def start_cv_process(message: Message, state: FSMContext):
    await message.answer("👤 Iltimos, ismingizni kiriting:", reply_markup=cancel_keyboard)
    await state.set_state(CVForm.waiting_for_name)


@dp.message(F.text == "❌ Bekor qilish")
async def cancel_cv_process(message: Message, state: FSMContext):
    current_state = await state.get_state()
    if current_state is not None:
        await state.clear()
        await message.answer("❌ CV jarayoni bekor qilindi.", reply_markup=main_menu)
    else:
        await message.answer("❗ Hozirda bekor qilinadigan jarayon yo‘q.", reply_markup=main_menu)


@dp.message(CVForm.waiting_for_name)
async def get_name(message: Message, state: FSMContext):
    if message.text == "❌ Bekor qilish":
        await cancel_cv_process(message, state)
        return
    await state.update_data(name=message.text)
    await message.answer("👥 Familiyangizni kiriting:", reply_markup=cancel_keyboard)
    await state.set_state(CVForm.waiting_for_surname)


@dp.message(CVForm.waiting_for_surname)
async def get_surname(message: Message, state: FSMContext):
    if message.text == "❌ Bekor qilish":
        await cancel_cv_process(message, state)
        return
    await state.update_data(surname=message.text)
    await message.answer("📞 Telefon raqamingizni kiriting (masalan: +998901234567):", reply_markup=cancel_keyboard)
    await state.set_state(CVForm.waiting_for_phone)


@dp.message(CVForm.waiting_for_phone)
async def get_phone_and_send_cv(message: Message, state: FSMContext):
    if message.text == "❌ Bekor qilish":
        await cancel_cv_process(message, state)
        return

    phone = message.text.strip()

    if not phone.startswith("+998") or len(phone) != 13:
        await message.answer("❌ Telefon raqam noto‘g‘ri formatda. Iltimos, +998 bilan kiriting:",
                             reply_markup=cancel_keyboard)
        return

    data = await state.get_data()
    name = data.get("name")
    surname = data.get("surname")
    logging.info(f"CV yuborildi: {name} {surname}, tel: {phone}, user: {message.from_user.id}")
    user_info = f"📥 CV so‘rov:\nIsm: {name}\nFamiliya: {surname}\nTelefon: {phone}\nUser: @{message.from_user.username or 'yo‘q'}"
    await bot.send_message(chat_id=int(ADMIN_ID), text=user_info)

    try:
        resume = FSInputFile("files/resume.pdf")
        await message.answer_document(resume, caption="📄 Mana mening CV faylim. Rahmat!", reply_markup=main_menu)
    except:
        await message.answer("❌ Kechirasiz, CV fayli topilmadi.", reply_markup=main_menu)

    await state.clear()


feedback_users = set()
feedback_storage = {}
feedback_counter = 0


class FeedbackReply(StatesGroup):
    waiting_for_reply = State()


@dp.message(lambda message: message.text == "✉️ Fikr bildirish")
async def ask_feedback(message: types.Message):
    feedback_users.add(message.from_user.id)
    await message.answer("✍️ Fikringizni yozing. Uni shaxsan adminga yuboraman.")


@dp.message(FeedbackReply.waiting_for_reply)
async def process_reply(message: Message, state: FSMContext):
    if str(message.from_user.id) != ADMIN_ID:
        await message.answer("❌ Bu buyruq faqat admin uchun!")
        await state.clear()
        return

    data = await state.get_data()
    feedback_id = data.get("feedback_id")

    if feedback_id is None or int(feedback_id) not in feedback_storage:
        await message.answer("❌ Bunday fikr ID topilmadi!")
        await state.clear()
        return

    feedback_id = int(feedback_id)
    user_id = feedback_storage[feedback_id]["user_id"]
    reply_text = message.text

    try:
        await bot.send_message(chat_id=user_id, text=f"📩 Admin javobi: {reply_text}")
        await message.answer(f"✅ Fikr ID {feedback_id} ga javob yuborildi.", reply_markup=main_menu)
    except Exception as e:
        await message.answer(f"❌ Xatolik yuz berdi: {str(e)}")
    finally:
        await state.clear()


@dp.message()
async def buttons(message: types.Message, state: FSMContext):
    global feedback_counter
    current_state = await state.get_state()
    if current_state is not None:
        await message.answer("❗ Iltimos, avval CV jarayonini yakunlang yoki bekor qiling.",
                             reply_markup=cancel_keyboard)
        return

    if message.from_user.id in feedback_users:
        feedback_users.discard(message.from_user.id)

        user = message.from_user
        feedback_counter += 1
        feedback_storage[feedback_counter] = {
            "user_id": user.id,
            "username": user.username or "username yo‘q",
            "text": message.text
        }

        reply_keyboard = InlineKeyboardMarkup(
            inline_keyboard=[
                [InlineKeyboardButton(text="✍️ Javob berish", callback_data=f"reply_{feedback_counter}")]
            ]
        )

        text = (
            f"✉️ <b>Yangi fikr keldi (ID: {feedback_counter})</b>\n\n"
            f"<b>Fikr:</b> {message.text}\n"
            f"<b>Foydalanuvchi:</b> @{user.username or 'username yo‘q'}\n"
            f"<b>User ID:</b> {user.id}"
        )

        await bot.send_message(chat_id=int(ADMIN_ID), text=text, reply_markup=reply_keyboard)
        await message.answer("✅ Fikringiz uchun rahmat!", reply_markup=main_menu)
    if message.text == "About":
        await message.answer("""
            👨‍💻 Men Bekmurod Maxmudov.  
        Toshkent Axborot Texnologiyalari Universitetining Farg‘ona filialida Kompyuter injiniringi yo‘nalishida tahsil olaman.

        🧠 Qiziqishlarim:
        - Python backend dasturlash
        - Telegram botlar yaratish
        - NoCode ilovalar (Adalo, FlutterFlow)
        - AI loyihalar

        🎯 Maqsadim — zamonaviy texnologiyalar orqali foydali loyihalar yaratish va ularni jamiyatga tatbiq etish.
            """)
    elif message.text == "Projects":
        await message.answer("""
           🛠 Mening ba’zi loyihalarim:

           1️⃣ **KinoBot** — Telegram orqali kinolar haqida ma'lumot beruvchi bot. (Rasm, kategoriya, trailer bilan)
           2️⃣ **Yordam+** — Tez yordam chaqirish va AI psixologik maslahat beruvchi mobil ilova (FlutterFlow + OpenAI API)
           3️⃣ **TarjimonBot** — PDF va Word fayllardan matn ajratib, 3 tilda tarjima qiluvchi Telegram bot
           4️⃣ **O‘quv markazi CRM** — Flask va SQLite3 asosida darslar, o‘quvchilar, to‘lovlar boshqaruvi tizimi
           """, reply_markup=projects_keyboard)

    elif message.text == "Skils":
        await message.answer("""
            🧰 Menda mavjud texnik ko‘nikmalar:

            👨‍💻 Dasturlash:
            - Python, HTML, CSS
            - Flask, aiogram
            - SQL (SQLite, MySQL)

            📱 Ilovalar:
            - Adalo, FlutterFlow (NoCode)
            - Telegram Bot API
            - Google Sheets API

            🧠 Boshqa:
            - Git & GitHub
            - Canva, Figma, PowerPoint
            - Matematika, statistika asoslari
            """)
    elif message.text == "Contact":
        await message.answer("📞 Mening kontaktlarim:", reply_markup=contact_keyboard)

    elif message.text == "Wikipedia":
        await message.answer(
            "Salom, Wikipedia Botiga xush kelibsiz! \n"
            "Botdan foydalanish uchun qidirmoqchi bo'lgan jumlangizni kiriting! \n"
            "Masalan:Wikipedia: 'O'zbekiston'"
        )
    elif message.text.startswith("Wikipedia:"):
        query = message.text.replace("Wikipedia:", "").strip()
        try:
            response = wikipedia.summary(query)
            await message.answer(response)
        except:
            await message.answer("❌ Bunday maqola mavjud emas!")


@dp.callback_query()
async def handle_callbacks(callback: CallbackQuery, state: FSMContext):
    data = callback.data
    if data.startswith("reply_"):
        feedback_id = int(data.replace("reply_", ""))
        if feedback_id in feedback_storage:
            await state.update_data(feedback_id=feedback_id)
            await callback.message.answer(
                f"📝 Fikr ID {feedback_id} ga javob yozing:",
                reply_markup=cancel_keyboard
            )
            await state.set_state(FeedbackReply.waiting_for_reply)
        else:
            await callback.message.answer("❌ Bunday fikr ID topilmadi.")
        await callback.answer()
        return

    if data == "KinoBot":
        await callback.message.answer("""
        💬 1. Kino izlash boti – @Kinokod11_bot
        🎬 Foydalanuvchilar kinoni nomi orqali izlaydi
        📦 JSON bazasidan ma’lumotlarni chiqaradi
        🔍 Inline mode qo‘llab-quvvatlanadi
        🧠 Eng ko‘p izlanuvchi kinolar bo‘yicha statistikasi bor
        🧪 Texnologiyalar: Python, Flask, SQLite, Telegram API
        """)
    elif data == "Yordam+":
        await    callback.message.answer("""
        📚 2. Yordam+ ilovasi – FlutterFlow + OpenAI API
        🚑 Tez yordam chaqirish funksiyasi
        🧠 AI asosida psixologik maslahatlar
        📱 NoCode bilan qurilgan UI/UX
        🔐 Ro‘yxatdan o‘tish, geolokatsiya bilan ishlash
        """)
    elif data == "TarjimonBot":
        await callback.message.answer("""
        🧾 3. Hujjat tarjima boti – @tarjimon_bot\n
        🌍 Ruscha, Inglizcha, O‘zbekcha tillar orasida tarjima\n
        📄 PDF va Word hujjatlar yuklab olinadi\n
        📌 Fayldan matnni ajratish va tarjima qilish\n
        🧪 Texnologiyalar: python-docx, pdfminer, Google Translate API\n
        """)
    else:
        await callback.message.answer("❗Nomaʼlum loyiha!")
    await callback.answer()


async def main():
    print("Bot ishaga tushdi !!!")
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
