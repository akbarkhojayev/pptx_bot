import asyncio
import logging
import os
import random
import textwrap
from io import BytesIO
from typing import List, Tuple, Optional
import wikipedia
from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import CommandStart
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.context import FSMContext
from aiogram.types import FSInputFile
from openai import OpenAI
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

BOT_TOKEN = ""
OPENAI_API_KEY = ""

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()
client = OpenAI(api_key=OPENAI_API_KEY)
logging.basicConfig(level=logging.INFO)
wikipedia.set_lang("uz")

class Form(StatesGroup):
    waiting_for_topic = State()
    waiting_for_image = State()
    waiting_for_type = State()
    waiting_for_more = State()

BACKGROUND_COLORS = [
    RGBColor(33, 150, 243),
    RGBColor(76, 175, 80),
    RGBColor(255, 152, 0),
    RGBColor(156, 39, 176),
    RGBColor(0, 188, 212),
]

# ---------------- PRESENTATION CREATOR ----------------
def split_text_to_chunks(text: str, max_chars: int = 1200) -> List[str]:
    """
    Split a given text into chunks of at most max_chars characters.
    """
    import re
    sentences = re.split(r"(?<=[.!?])\s+", text.strip())
    chunks, current = [], ""
    for s in sentences:
        if len(current) + len(s) + 1 <= max_chars:
            current = (current + " " + s).strip()
        else:
            if current:
                chunks.append(current.strip())
            if len(s) > max_chars:
                chunks.extend(textwrap.wrap(s, max_chars))
                current = ""
            else:
                current = s
    if current:
        chunks.append(current.strip())
    return chunks


# ---------------- PRESENTATION CREATOR ----------------
def create_presentation_file(topic: str, plan: List[str], content_chunks: List[str], user_id: int, slide_images: list, large_images: list) -> str:
    prs = Presentation()
    prs.slide_height = Inches(7.5)
    prs.slide_width = Inches(13.3333)

    # Title + Plan slide
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = RGBColor(25, 118, 210)

    title_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.5), prs.slide_width - Inches(1.2), Inches(1.2))
    title_tf = title_box.text_frame
    title_p = title_tf.paragraphs[0]
    title_p.text = topic
    title_p.font.size = Pt(44)
    title_p.font.bold = True
    title_p.font.color.rgb = RGBColor(255, 255, 255)
    title_p.alignment = PP_ALIGN.CENTER

    sub_box = slide.shapes.add_textbox(Inches(1.0), Inches(1.9), prs.slide_width - Inches(2.0), Inches(0.6))
    sub_tf = sub_box.text_frame
    sp = sub_tf.paragraphs[0]
    sp.text = "Reja"
    sp.font.size = Pt(28)
    sp.font.bold = True
    sp.font.color.rgb = RGBColor(255, 255, 255)
    sp.alignment = PP_ALIGN.LEFT

    bullets_box = slide.shapes.add_textbox(Inches(1.0), Inches(2.6), prs.slide_width - Inches(2.0), Inches(4.0))
    bullets_tf = bullets_box.text_frame
    bullets_tf.word_wrap = True
    for item in plan:
        p = bullets_tf.add_paragraph()
        p.text = f"• {item}"
        p.font.size = Pt(18)
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.alignment = PP_ALIGN.LEFT

    # Content slides: left = text, right = image
    left_margin = Inches(0.6)
    right_margin = Inches(0.6)
    gap = Inches(0.3)
    image_width = Inches(4.0)
    image_height = Inches(4.0)
    text_top = Inches(1.2)
    text_height = Inches(5.0)

    # ensure at least 1 image per content slide
    needed = max(7, len(content_chunks))
    chunks = content_chunks.copy()
    # if too few chunks, split long ones
    if len(chunks) < needed:
        idx = 0
        while len(chunks) < needed and idx < len(chunks):
            if len(chunks[idx]) > 800:
                extra = split_text_to_chunks(chunks[idx], max_chars=700)
                # replace this chunk with its parts
                chunks.pop(idx)
                for j, part in enumerate(extra):
                    chunks.insert(idx + j, part)
            idx += 1
            if idx >= len(chunks):
                break
    final_chunks = chunks[:needed]

    # Ensure minimum text per slide
    minimum_chars = 400
    i = 0
    while i < len(final_chunks) - 1:
        if len(final_chunks[i]) < minimum_chars:
            # Merge with next
            final_chunks[i] += " " + final_chunks[i+1]
            final_chunks.pop(i+1)
        else:
            i += 1
    # If last is too short and more than one, merge with previous
    if len(final_chunks) > 1 and len(final_chunks[-1]) < minimum_chars:
        final_chunks[-2] += " " + final_chunks[-1]
        final_chunks.pop()

    slide_num = 1
    for idx, chunk in enumerate(final_chunks):
        if large_images and idx == len(final_chunks) // 2:
            # Add large image slide in the middle
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = RGBColor(255, 255, 255)
            # Add large image
            img_path = large_images[0]  # Use the first large image
            try:
                with open(img_path, "rb") as f:
                    img_data = f.read()
                img_stream = BytesIO(img_data)
                img_stream.seek(0)
                slide.shapes.add_picture(img_stream, Inches(0.5), Inches(0.5), width=prs.slide_width - Inches(1), height=prs.slide_height - Inches(1))
            except Exception as e:
                logging.exception(f"Failed to add large image: {e}")

        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(255, 255, 255)

        # Heading
        title_box = slide.shapes.add_textbox(left_margin, Inches(0.5), prs.slide_width - left_margin - right_margin, Inches(1.0))
        tf = title_box.text_frame
        p = tf.paragraphs[0]
        heading = f"{slide_num}. {plan[idx] if idx < len(plan) else topic}"
        p.text = heading
        p.font.size = Pt(24)
        p.font.bold = True
        p.font.color.rgb = RGBColor(0, 0, 0)
        p.alignment = PP_ALIGN.LEFT

        # Text
        if slide_images and idx < len(slide_images):
            # With image: left column
            content_box_width = prs.slide_width - left_margin - right_margin - image_width - gap
            content_box = slide.shapes.add_textbox(left_margin, text_top, content_box_width, text_height)
            content_tf = content_box.text_frame
            content_tf.word_wrap = True
            for part in split_text_to_chunks(chunk, max_chars=600):
                para = content_tf.add_paragraph()
                para.text = part
                para.font.size = Pt(16)
                para.font.name = "Segoe UI"
                para.font.color.rgb = RGBColor(0, 0, 0)
                para.alignment = PP_ALIGN.JUSTIFY

            # Image (right column)
            img_path = slide_images[idx]
            try:
                with open(img_path, "rb") as f:
                    img_bytes = f.read()
                img_stream = BytesIO(img_bytes)
                img_stream.seek(0)
                slide.shapes.add_picture(img_stream, left_margin + content_box_width + gap, text_top, width=image_width, height=image_height)
            except Exception as e:
                logging.exception(f"Failed to add picture on slide {idx+1}: {e}")
        else:
            # No image: full width, centered
            content_box = slide.shapes.add_textbox(left_margin, text_top, prs.slide_width - left_margin - right_margin, text_height)
            content_tf = content_box.text_frame
            content_tf.word_wrap = True
            for part in split_text_to_chunks(chunk, max_chars=800):  # More chars since full width
                para = content_tf.add_paragraph()
                para.text = part
                para.font.size = Pt(18)  # Slightly larger
                para.font.name = "Segoe UI"
                para.font.color.rgb = RGBColor(0, 0, 0)
                para.alignment = PP_ALIGN.JUSTIFY

        slide_num += 1

    # Final "Thank you" slide
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = RGBColor(255, 255, 255)
    box = slide.shapes.add_textbox(Inches(1), Inches(2.5), prs.slide_width - Inches(2), Inches(3))
    tf = box.text_frame
    p = tf.paragraphs[0]
    p.text = "ETIBORINGIZ UCHUN RAHMAT!"
    p.font.size = Pt(48)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0, 0, 0)
    p.alignment = PP_ALIGN.CENTER

    filename = f"{topic.replace(' ', '_')}.pptx"
    prs.save(filename)
    return filename


# ---------------- OPENAI TEXT GENERATION ----------------
def generate_plan_and_contents(topic: str) -> Tuple[List[str], List[str]]:
    try:
        wiki = wikipedia.summary(topic, sentences=6)
    except Exception:
        wiki = ""

    user_prompt = f"""
Mavzu: {topic}
Wikipedia: {wiki}

Vazifa:
1) 4-5 banddan iborat reja tuzing (raqamlangan), 5 - bu xulosa bolsin.
2) Har band uchun 300-500 so'zli tushunarli matn yozing.
3) Matn misollar bilan bo'lsin va oxirida qisqacha xulosa yozing.

Natija quyidagi formatda bo'lsin:
Reja:
1. ...
2. ...
Matnlar:
1. ...
2. ...
"""

    try:
        resp = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "Siz professional taqdimot yaratuvchisiz."},
                {"role": "user", "content": user_prompt},
            ],
            temperature=0.6,
            max_tokens=3000,
        )
        raw = resp.choices[0].message.content.strip()
        # Clean unnecessary characters
        import re
        raw = re.sub(r'[*#]', '', raw)
    except Exception as e:
        logging.exception("OpenAI xatosi:")
        # fallback simple content
        plan = ["Kirish", "Asosiy tushunchalar", "Amaliy misollar", "Muammolar va echimlar", "Xulosa"]
        contents = ["Kirish: mavzu haqida umumiy ma'lumot."] * 5
        return plan, contents

    plan = []
    contents = []
    if "Reja:" in raw and "Matnlar:" in raw:
        try:
            plan_part = raw.split("Matnlar:")[0].replace("Reja:", "").strip()
            text_part = raw.split("Matnlar:")[1].strip()
            for line in plan_part.splitlines():
                line = line.strip()
                if line and (line[0].isdigit() or line.startswith("-")):
                    if "." in line:
                        plan.append(line.split(".", 1)[1].strip())
                    else:
                        plan.append(line.lstrip("- ").strip())
            cur_idx = 0
            current_texts = {}
            for line in text_part.splitlines():
                line = line.strip()
                if not line:
                    continue
                if line[0].isdigit() and "." in line[:3]:
                    idx = int(line.split(".", 1)[0].strip())
                    rest = line.split(".", 1)[1].strip()
                    current_texts[idx] = rest
                    cur_idx = idx
                else:
                    if cur_idx == 0:
                        current_texts.setdefault(1, "")
                        current_texts[1] += " " + line
                    else:
                        current_texts[cur_idx] = current_texts.get(cur_idx, "") + " " + line
            max_idx = max(current_texts.keys()) if current_texts else 0
            for i in range(1, max_idx + 1):
                contents.append(current_texts.get(i, "").strip())
        except Exception:
            plan = ["Kirish", "Asosiy tushunchalar", "Amaliy misollar", "Muammolar va echimlar", "Xulosa"]
            contents = ["Kirish: mavzu haqida umumiy ma'lumot."] * 5
    else:
        # fallback simple split
        lines = [ln.strip() for ln in raw.splitlines() if ln.strip()]
        plan = lines[:5]
        contents = [" ".join(lines[5:]) or "Ma'lumot yetishmadi."] * len(plan)

    # ensure equal lengths
    while len(contents) < len(plan):
        contents.append("Ma'lumot yetishmadi — iltimos mavzuni kengroq yozing.")

    return plan, contents


# ---------------- BOT HANDLERS ----------------
@dp.message(CommandStart())
async def cmd_start(message: types.Message, state: FSMContext):
    await state.clear()
    await message.answer(
        "👋 Salom! Men AI Prezentatsiya Botman.\n\n"
        "Prezentatsiya yaratish uchun mavzu kiriting:"
    )
    await state.set_state(Form.waiting_for_topic)


@dp.message(F.photo, Form.waiting_for_more)
async def handle_additional_photo(message: types.Message, state: FSMContext):
    photo = message.photo[-1]
    file = await bot.get_file(photo.file_id)
    data = await state.get_data()
    img_index = len(data.get('slide_images', [])) + len(data.get('large_images', []))
    user_img_path = f"user_image_{message.from_user.id}_{img_index}.png"
    await bot.download_file(file.file_path, user_img_path)
    await state.update_data(last_img_path=user_img_path)
    await message.reply("🖼 Rasm yuklandi! Bu rasmni qayerda ishlatmoqchisiz?\n- 'asosiy' - Katta sahifada\n- 'matn' - Matn sahifalarida")
    await state.set_state(Form.waiting_for_type)


@dp.message(F.photo, Form.waiting_for_image)
async def handle_photo(message: types.Message, state: FSMContext):
    photo = message.photo[-1]
    file = await bot.get_file(photo.file_id)
    data = await state.get_data()
    img_index = len(data.get('slide_images', [])) + len(data.get('large_images', []))
    user_img_path = f"user_image_{message.from_user.id}_{img_index}.png"
    await bot.download_file(file.file_path, user_img_path)
    await state.update_data(last_img_path=user_img_path)
    await message.reply("🖼 Rasm yuklandi! Bu rasmni qayerda ishlatmoqchisiz?\n- 'asosiy' - Katta sahifada\n- 'matn' - Matn sahifalarida")
    await state.set_state(Form.waiting_for_type)


@dp.message(Form.waiting_for_type)
async def handle_type(message: types.Message, state: FSMContext):
    data = await state.get_data()
    img_path = data.get('last_img_path')
    if not img_path:
        await message.reply("Xatolik yuz berdi. Qaytadan boshlang.")
        await state.clear()
        return
    typ = message.text.lower().strip()
    if typ == 'asosiy':
        large_images = data.get('large_images', [])
        large_images.append(img_path)
        await state.update_data(large_images=large_images)
    elif typ == 'matn':
        slide_images = data.get('slide_images', [])
        slide_images.append(img_path)
        await state.update_data(slide_images=slide_images)
    else:
        await state.clear()
        await message.reply("Noto'g'ri javob. Qaytadan boshlang: /start")
        return
    await message.reply("✅ Rasm saqlandi! Yana rasm yuklamoqchimisiz? 'done' yozing")
    await state.set_state(Form.waiting_for_more)
@dp.message(Form.waiting_for_topic)
async def handle_topic_input(message: types.Message, state: FSMContext):
    topic = message.text.strip()
    if not topic:
        await message.reply("Iltimos, mavzuni yozing.")
        return
    await state.update_data(topic=topic, slide_images=[], large_images=[])
    await message.answer("✅ Mavzu saqlandi.\n\nEndi prezentatsiya uchun rasm yuboring yoki 'skip' yozing (rasmsiz yaratish uchun):")
    await state.set_state(Form.waiting_for_image)


@dp.message(Form.waiting_for_more)
async def handle_more(message: types.Message, state: FSMContext):
    if message.text and message.text.lower().strip() == "done":
        data = await state.get_data()
        topic = data.get("topic")
        slide_images = data.get('slide_images', [])
        large_images = data.get('large_images', [])
        await state.clear()
        await message.reply("Prezentatsiya yaratilmoqda...")
        await create_presentation(message, topic, slide_images, large_images)
    else:
        await message.reply("Yana rasm yuklamoqchimisiz? 'done' yozing")


@dp.message()
async def handle_unknown(message: types.Message):
    await message.reply("Boshlash uchun /start ni bosing.")
async def create_presentation(message: types.Message, topic: str, slide_images: list, large_images: list):
    await message.answer(f"🔍 '{topic}' bo'yicha ma'lumotlar olinmoqda...")
    loop = asyncio.get_event_loop()
    plan, contents = await loop.run_in_executor(None, generate_plan_and_contents, topic)

    content_chunks = []
    for text in contents:
        content_chunks.extend(split_text_to_chunks(text, max_chars=800))

    await message.answer("🎨 Dizayn va rasm tanlanmoqda, fayl yaratilmoqda...")
    pptx_path = await loop.run_in_executor(None, create_presentation_file, topic, plan, content_chunks, message.from_user.id, slide_images, large_images)

    await message.answer("✅ Prezentatsiya tayyor! Yuklab olayotganman...")
    await message.answer_document(FSInputFile(pptx_path))

    try:
        os.remove(pptx_path)
    except Exception:
        pass

    # Clean up user images after use
    all_images = slide_images + large_images
    for img_path in all_images:
        try:
            if os.path.exists(img_path):
                os.remove(img_path)
        except Exception:
            pass


if __name__ == "__main__":
    logging.info("🚀 Bot ishga tushmoqda...")
    asyncio.run(dp.start_polling(bot))
