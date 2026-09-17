import asyncio
import logging
import math
import os
import random
import re
import textwrap
from io import BytesIO
from typing import List, Tuple, Optional
import wikipedia
import wikipedia.wikipedia as _wikipedia_internal
from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import CommandStart
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.context import FSMContext
from aiogram.types import FSInputFile
from openai import OpenAI
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt

BOT_TOKEN = ""

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()
logging.basicConfig(level=logging.INFO)
wikipedia.set_lang("uz")
# Wikimedia rate-limits the library's shared default User-Agent; a distinct one avoids collateral 429s.
_wikipedia_internal.USER_AGENT = "PptxPresentationBot/1.0 (Telegram content-generation bot)"

class Form(StatesGroup):
    waiting_for_topic = State()
    waiting_for_image = State()
    waiting_for_type = State()
    waiting_for_more = State()

# ---------------- DESIGN THEMES ----------------
FONT_HEADING = "Century Gothic"
FONT_BODY = "Calibri"

THEMES = [
    {
        "name": "Ocean",
        "primary": RGBColor(0x0D, 0x47, 0xA1),
        "primary_dark": RGBColor(0x07, 0x2A, 0x5E),
        "accent": RGBColor(0x29, 0xB6, 0xF6),
        "bg": RGBColor(0xF5, 0xF9, 0xFF),
        "text": RGBColor(0x16, 0x22, 0x33),
        "muted": RGBColor(0x7A, 0x8B, 0xA3),
    },
    {
        "name": "Emerald",
        "primary": RGBColor(0x1B, 0x5E, 0x20),
        "primary_dark": RGBColor(0x0E, 0x35, 0x12),
        "accent": RGBColor(0x66, 0xBB, 0x6A),
        "bg": RGBColor(0xF3, 0xFA, 0xF3),
        "text": RGBColor(0x17, 0x26, 0x18),
        "muted": RGBColor(0x7C, 0x94, 0x7E),
    },
    {
        "name": "Sunset",
        "primary": RGBColor(0xBF, 0x36, 0x0C),
        "primary_dark": RGBColor(0x6E, 0x1E, 0x08),
        "accent": RGBColor(0xFF, 0xA7, 0x26),
        "bg": RGBColor(0xFF, 0xF8, 0xF1),
        "text": RGBColor(0x2E, 0x1C, 0x12),
        "muted": RGBColor(0xA3, 0x87, 0x74),
    },
    {
        "name": "Royal",
        "primary": RGBColor(0x4A, 0x14, 0x8C),
        "primary_dark": RGBColor(0x28, 0x0A, 0x4D),
        "accent": RGBColor(0xBA, 0x68, 0xC8),
        "bg": RGBColor(0xF9, 0xF5, 0xFC),
        "text": RGBColor(0x22, 0x16, 0x2C),
        "muted": RGBColor(0x93, 0x82, 0xA0),
    },
    {
        "name": "Slate",
        "primary": RGBColor(0x1C, 0x2B, 0x36),
        "primary_dark": RGBColor(0x0A, 0x14, 0x1A),
        "accent": RGBColor(0x26, 0xC6, 0xDA),
        "bg": RGBColor(0xF4, 0xF7, 0xF8),
        "text": RGBColor(0x14, 0x1D, 0x22),
        "muted": RGBColor(0x7C, 0x92, 0x9C),
    },
]


def lighten(color: RGBColor, factor: float) -> RGBColor:
    r, g, b = color[0], color[1], color[2]
    return RGBColor(
        int(r + (255 - r) * factor),
        int(g + (255 - g) * factor),
        int(b + (255 - b) * factor),
    )


def set_shape_transparency(shape, alpha_pct: int) -> None:
    """alpha_pct: 0 (invisible) .. 100 (fully opaque)."""
    solid_fill = shape.fill.fore_color._xFill
    srgb_clr = solid_fill.find(qn("a:srgbClr"))
    alpha = OxmlElement("a:alpha")
    alpha.set("val", str(int(alpha_pct * 1000)))
    srgb_clr.append(alpha)


def no_shadow(shape) -> None:
    try:
        shape.shadow.inherit = False
    except Exception:
        pass


def add_gradient_bg(slide, color1: RGBColor, color2: RGBColor, angle: float = 45) -> None:
    fill = slide.background.fill
    fill.gradient()
    stops = fill.gradient_stops
    stops[0].color.rgb = color1
    stops[1].color.rgb = color2
    fill.gradient_angle = angle


def add_decor_circle(slide, x, y, d, color: RGBColor, alpha_pct: int = 15):
    shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, d, d)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    no_shadow(shape)
    set_shape_transparency(shape, alpha_pct)
    return shape


def add_rect(slide, x, y, w, h, color: RGBColor, shape_type=MSO_SHAPE.RECTANGLE):
    shape = slide.shapes.add_shape(shape_type, x, y, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    no_shadow(shape)
    return shape


def add_footer(slide, theme: dict, topic: str, page_no: int, prs_width) -> None:
    add_rect(slide, Inches(0.6), Inches(6.95), prs_width - Inches(1.2), Pt(1.2), lighten(theme["primary"], 0.6))
    box = slide.shapes.add_textbox(Inches(0.6), Inches(7.02), prs_width - Inches(1.6), Inches(0.35))
    tf = box.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.text = topic
    p.font.size = Pt(11)
    p.font.name = FONT_BODY
    p.font.color.rgb = theme["muted"]
    p.alignment = PP_ALIGN.LEFT

    num_box = slide.shapes.add_textbox(prs_width - Inches(1.2), Inches(7.02), Inches(0.6), Inches(0.35))
    ntf = num_box.text_frame
    np_ = ntf.paragraphs[0]
    np_.text = str(page_no)
    np_.font.size = Pt(11)
    np_.font.name = FONT_BODY
    np_.font.color.rgb = theme["muted"]
    np_.alignment = PP_ALIGN.RIGHT


def _text_block_height_pt(paragraphs: List[str], size: int, box_w_pt: float, line_spacing: float, space_after_pt: float) -> float:
    """Height (pt) needed to lay out `paragraphs` at font `size` in a box `box_w_pt` wide.
    Calibrated against real LibreOffice-rendered glyph positions (see project notes):
    average glyph width ~0.5x font size, and line height is ~1.2x font size (base leading)
    times the paragraph's line_spacing multiplier -- e.g. at size=19/line_spacing=1.2 this
    predicts 27.4pt between lines, matching the 27.3pt actually measured."""
    avg_char_w_pt = size * 0.5
    chars_per_line = max(1, box_w_pt / avg_char_w_pt)
    line_h_pt = size * 1.2 * line_spacing
    total_lines = sum(max(1, math.ceil(len(p) / chars_per_line)) for p in paragraphs)
    return total_lines * line_h_pt + max(0, len(paragraphs) - 1) * space_after_pt


def pick_font_size(
    paragraphs: List[str],
    box_w_emu: int,
    box_h_emu: int,
    sizes: Tuple[int, ...],
    line_spacing: float = 1.2,
    space_after_fn=lambda size: max(6, size // 2),
) -> int:
    """Pick the largest font size (pt) from `sizes` whose paragraphs fit box_h_emu tall,
    so slide text never overflows its box regardless of how long the source content is."""
    box_w_pt = box_w_emu / 12700
    box_h_pt = box_h_emu / 12700 * 0.95  # small safety margin
    for size in sizes:
        needed = _text_block_height_pt(paragraphs, size, box_w_pt, line_spacing, space_after_fn(size))
        if needed <= box_h_pt:
            return size
    return sizes[-1]


def fit_dimensions(img_path: str, max_w: int, max_h: int) -> Tuple[int, int]:
    """Return (width, height) in EMU that fit inside max_w x max_h, preserving aspect ratio."""
    try:
        with Image.open(img_path) as im:
            iw, ih = im.size
        ratio = min(max_w / iw, max_h / ih)
        return int(iw * ratio), int(ih * ratio)
    except Exception:
        return max_w, max_h

# ---------------- PRESENTATION CREATOR ----------------
def split_text_to_chunks(text: str, max_chars: int = 1200) -> List[str]:
    """
    Split a given text into chunks of at most max_chars characters.
    """
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
    sw, sh = prs.slide_width, prs.slide_height

    theme = random.choice(THEMES)

    # ---------------- Title + Plan slide ----------------
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_gradient_bg(slide, theme["primary_dark"], theme["primary"], angle=45)

    add_decor_circle(slide, sw - Inches(2.6), Inches(-2.0), Inches(5.2), RGBColor(255, 255, 255), alpha_pct=8)
    # (no bottom-left decoration here: the plan list can grow to ~9 items and fill that corner,
    # and a background circle behind the text created a patchy, "broken"-looking backdrop)

    title_box = slide.shapes.add_textbox(Inches(0.9), Inches(0.7), sw - Inches(1.8), Inches(1.5))
    title_tf = title_box.text_frame
    title_tf.word_wrap = True
    title_p = title_tf.paragraphs[0]
    title_p.text = topic
    title_p.font.size = Pt(42)
    title_p.font.bold = True
    title_p.font.name = FONT_HEADING
    title_p.font.color.rgb = RGBColor(255, 255, 255)
    title_p.alignment = PP_ALIGN.LEFT

    add_rect(slide, Inches(0.95), Inches(1.95), Inches(1.4), Pt(4), theme["accent"])

    badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.9), Inches(2.3), Inches(1.5), Inches(0.5))
    badge.fill.solid()
    badge.fill.fore_color.rgb = theme["accent"]
    badge.line.fill.background()
    no_shadow(badge)
    btf = badge.text_frame
    btf.vertical_anchor = MSO_ANCHOR.MIDDLE
    bp = btf.paragraphs[0]
    bp.text = "REJA"
    bp.font.size = Pt(16)
    bp.font.bold = True
    bp.font.name = FONT_BODY
    bp.font.color.rgb = RGBColor(255, 255, 255)
    bp.alignment = PP_ALIGN.CENTER

    n_items = max(len(plan), 1)
    list_top = Inches(3.15)
    list_bottom = sh - Inches(0.35)
    item_h = min(Inches(0.72), int((list_bottom - list_top) / n_items))
    badge_d = min(Inches(0.42), max(Inches(0.26), item_h - Inches(0.14)))
    if item_h >= Inches(0.6):
        item_font, num_font = Pt(19), Pt(16)
    elif item_h >= Inches(0.45):
        item_font, num_font = Pt(15), Pt(13)
    else:
        item_font, num_font = Pt(12), Pt(11)

    item_y = list_top
    for i, item in enumerate(plan, start=1):
        badge_y = item_y + (item_h - badge_d) // 2
        num = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.9), badge_y, badge_d, badge_d)
        num.fill.solid()
        num.fill.fore_color.rgb = theme["accent"]
        num.line.color.rgb = RGBColor(255, 255, 255)
        num.line.width = Pt(1.25)
        no_shadow(num)
        ntf = num.text_frame
        ntf.vertical_anchor = MSO_ANCHOR.MIDDLE
        np_ = ntf.paragraphs[0]
        np_.text = str(i)
        np_.font.size = num_font
        np_.font.bold = True
        np_.font.name = FONT_BODY
        np_.font.color.rgb = RGBColor(255, 255, 255)
        np_.alignment = PP_ALIGN.CENTER

        text_box = slide.shapes.add_textbox(Inches(1.55), item_y, sw - Inches(2.6), item_h)
        ttf = text_box.text_frame
        ttf.word_wrap = True
        ttf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tp = ttf.paragraphs[0]
        tp.text = item
        tp.font.size = item_font
        tp.font.name = FONT_BODY
        tp.font.color.rgb = RGBColor(255, 255, 255)
        tp.alignment = PP_ALIGN.LEFT
        item_y += item_h

    # ---------------- Content slide layout constants ----------------
    left_margin = Inches(0.7)
    right_margin = Inches(0.7)
    gap = Inches(0.4)
    image_box_w = Inches(4.3)
    image_box_h = Inches(4.6)
    text_top = Inches(1.55)
    text_height = Inches(5.15)

    # guarantee >=10 total slides: 1 title + needed content + 1 thank-you (+1 if a large image is inserted)
    needed = max(8, len(content_chunks))
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
            # ---------------- Large image slide ----------------
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = theme["bg"]
            add_rect(slide, 0, 0, sw, Inches(0.14), theme["accent"])

            img_path = large_images[0]  # Use the first large image
            try:
                box_x, box_y, box_w, box_h = Inches(0.9), Inches(0.7), sw - Inches(1.8), sh - Inches(1.9)
                pic_w, pic_h = fit_dimensions(img_path, box_w, box_h)
                pic_x = box_x + (box_w - pic_w) // 2
                pic_y = box_y + (box_h - pic_h) // 2
                frame_pad = Inches(0.08)
                add_rect(slide, pic_x - frame_pad, pic_y - frame_pad, pic_w + frame_pad * 2, pic_h + frame_pad * 2, theme["primary"])
                with open(img_path, "rb") as f:
                    img_data = f.read()
                img_stream = BytesIO(img_data)
                img_stream.seek(0)
                slide.shapes.add_picture(img_stream, pic_x, pic_y, width=pic_w, height=pic_h)

                caption = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, sw / 2 - Inches(2.5), sh - Inches(1.05), Inches(5.0), Inches(0.55))
                caption.fill.solid()
                caption.fill.fore_color.rgb = theme["primary"]
                caption.line.fill.background()
                no_shadow(caption)
                ctf = caption.text_frame
                ctf.vertical_anchor = MSO_ANCHOR.MIDDLE
                cp = ctf.paragraphs[0]
                cp.text = topic
                cp.font.size = Pt(16)
                cp.font.bold = True
                cp.font.name = FONT_BODY
                cp.font.color.rgb = RGBColor(255, 255, 255)
                cp.alignment = PP_ALIGN.CENTER
            except Exception as e:
                logging.exception(f"Failed to add large image: {e}")

        # ---------------- Regular content slide ----------------
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = theme["bg"]
        add_rect(slide, 0, 0, sw, Inches(0.14), theme["accent"])

        # Slide-number badge + heading
        badge_d2 = Inches(0.55)
        num_badge = slide.shapes.add_shape(MSO_SHAPE.OVAL, left_margin, Inches(0.55), badge_d2, badge_d2)
        num_badge.fill.solid()
        num_badge.fill.fore_color.rgb = theme["primary"]
        num_badge.line.fill.background()
        no_shadow(num_badge)
        nbtf = num_badge.text_frame
        nbtf.word_wrap = False
        nbtf.vertical_anchor = MSO_ANCHOR.MIDDLE
        nbp = nbtf.paragraphs[0]
        nbp.text = str(slide_num)
        nbp.font.size = Pt(20) if slide_num < 10 else Pt(15)
        nbp.font.bold = True
        nbp.font.name = FONT_BODY
        nbp.font.color.rgb = RGBColor(255, 255, 255)
        nbp.alignment = PP_ALIGN.CENTER

        title_box_w = sw - left_margin - right_margin - badge_d2 - Inches(0.25)
        title_box = slide.shapes.add_textbox(left_margin + badge_d2 + Inches(0.25), Inches(0.45), title_box_w, Inches(0.85))
        tf = title_box.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        heading = plan[idx] if idx < len(plan) else topic
        p.text = heading
        p.font.size = Pt(pick_font_size([heading], title_box_w, Inches(0.85), sizes=(26, 24, 22, 20, 18), line_spacing=1.0, space_after_fn=lambda s: 0))
        p.font.bold = True
        p.font.name = FONT_HEADING
        p.font.color.rgb = theme["primary"]
        p.alignment = PP_ALIGN.LEFT

        add_rect(slide, left_margin, Inches(1.3), Inches(0.9), Pt(3), theme["accent"])

        # Text
        if slide_images and idx < len(slide_images):
            # With image: left column, divider, image on right (aspect-ratio preserved)
            content_box_width = sw - left_margin - right_margin - image_box_w - gap
            content_box = slide.shapes.add_textbox(left_margin, text_top, content_box_width, text_height)
            content_tf = content_box.text_frame
            content_tf.word_wrap = True
            paragraphs = split_text_to_chunks(chunk, max_chars=600)
            body_size = pick_font_size(paragraphs, content_box_width, text_height, sizes=(17, 16, 15, 14, 13, 12, 11), line_spacing=1.15)
            for part in paragraphs:
                para = content_tf.add_paragraph()
                para.text = part
                para.font.size = Pt(body_size)
                para.font.name = FONT_BODY
                para.font.color.rgb = theme["text"]
                para.alignment = PP_ALIGN.JUSTIFY
                para.line_spacing = 1.15
                para.space_after = Pt(max(6, body_size // 2))

            add_rect(slide, left_margin + content_box_width + gap / 2, text_top, Pt(1.2), text_height, lighten(theme["primary"], 0.55))

            img_path = slide_images[idx]
            try:
                img_area_x = left_margin + content_box_width + gap
                pic_w, pic_h = fit_dimensions(img_path, image_box_w, image_box_h)
                pic_x = img_area_x + (image_box_w - pic_w) // 2
                pic_y = text_top + (image_box_h - pic_h) // 2
                frame_pad = Inches(0.06)
                add_rect(slide, pic_x - frame_pad, pic_y - frame_pad, pic_w + frame_pad * 2, pic_h + frame_pad * 2, lighten(theme["primary"], 0.85))
                with open(img_path, "rb") as f:
                    img_bytes = f.read()
                img_stream = BytesIO(img_bytes)
                img_stream.seek(0)
                pic = slide.shapes.add_picture(img_stream, pic_x, pic_y, width=pic_w, height=pic_h)
                pic.line.color.rgb = theme["accent"]
                pic.line.width = Pt(1.5)
            except Exception as e:
                logging.exception(f"Failed to add picture on slide {idx+1}: {e}")
        else:
            # No image: full width, centered
            full_width = sw - left_margin - right_margin
            content_box = slide.shapes.add_textbox(left_margin, text_top, full_width, text_height)
            content_tf = content_box.text_frame
            content_tf.word_wrap = True
            paragraphs = split_text_to_chunks(chunk, max_chars=800)  # More chars since full width
            body_size = pick_font_size(paragraphs, full_width, text_height, sizes=(19, 18, 17, 16, 15, 14, 13, 12), line_spacing=1.2)
            for part in paragraphs:
                para = content_tf.add_paragraph()
                para.text = part
                para.font.size = Pt(body_size)
                para.font.name = FONT_BODY
                para.font.color.rgb = theme["text"]
                para.alignment = PP_ALIGN.JUSTIFY
                para.line_spacing = 1.2
                para.space_after = Pt(max(6, body_size // 2))

        add_footer(slide, theme, topic, slide_num, sw)
        slide_num += 1

    # ---------------- Final "Thank you" slide ----------------
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_gradient_bg(slide, theme["primary"], theme["primary_dark"], angle=225)
    add_decor_circle(slide, sw - Inches(3.2), sh - Inches(3.2), Inches(4.8), RGBColor(255, 255, 255), alpha_pct=8)
    add_decor_circle(slide, Inches(-1.2), Inches(-1.2), Inches(3.2), theme["accent"], alpha_pct=18)

    box = slide.shapes.add_textbox(Inches(1), Inches(2.7), sw - Inches(2), Inches(1.3))
    tf = box.text_frame
    p = tf.paragraphs[0]
    p.text = "RAHMAT!"
    p.font.size = Pt(54)
    p.font.bold = True
    p.font.name = FONT_HEADING
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.alignment = PP_ALIGN.CENTER

    add_rect(slide, sw / 2 - Inches(0.75), Inches(3.75), Inches(1.5), Pt(4), theme["accent"])

    sub_box = slide.shapes.add_textbox(Inches(1), Inches(4.0), sw - Inches(2), Inches(0.7))
    stf = sub_box.text_frame
    sp = stf.paragraphs[0]
    sp.text = topic
    sp.font.size = Pt(20)
    sp.font.name = FONT_BODY
    sp.font.color.rgb = lighten(theme["accent"], 0.3)
    sp.alignment = PP_ALIGN.CENTER

    filename = f"{topic.replace(' ', '_')}.pptx"
    prs.save(filename)
    return filename


# ---------------- GROQ AI CONTENT GENERATION ----------------
GROQ_API_KEY = "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"
GROQ_MODEL = "openai/gpt-oss-120b"
MIN_SECTIONS = 9  # Kirish + 7 mavzu bandi + Xulosa => kamida 10 slaydga yetadi

groq_client = OpenAI(api_key=GROQ_API_KEY, base_url="https://api.groq.com/openai/v1")


def _parse_plan_response(raw: str) -> Tuple[List[str], List[str]]:
    plan: List[str] = []
    contents: List[str] = []
    if "Reja:" not in raw or "Matnlar:" not in raw:
        return plan, contents
    try:
        plan_part = raw.split("Matnlar:")[0].split("Reja:", 1)[1].strip()
        text_part = raw.split("Matnlar:", 1)[1].strip()
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
        return [], []
    return plan, contents


GROQ_SYSTEM_PROMPT = """Siz "PPT Yordamchi" — professional prezentatsiya strukturasi va kontent bo'yicha mutaxassissiz.

Har bir mavzu uchun avval uning turini o'zingiz aniqlang (ilmiy/diplom, biznes/startup yoki ta'lim/dars) va rejani o'sha turga mos tuzing:
- Ilmiy/diplom uslubidagi mavzular uchun: kirish, maqsad va vazifalar, nazariy asoslar, amaliy/tahliliy qism, natijalar, xulosa kabi ketma-ketlikka moslashtiring.
- Biznes/startup uslubidagi mavzular uchun: muammo, yechim, bozor/qo'llanilish sohasi, ustunliklar, amaliyot, istiqbollar kabi ketma-ketlikka moslashtiring.
- Umumiy ta'lim/dars mavzulari uchun: kirish, asosiy tushunchalar, tarixi yoki rivojlanishi, turlari/tasnifi, amaliy qo'llanilishi, muammo va yechimlar kabi ketma-ketlikka moslashtiring.

Har doim: aniq faktlar, raqamlar va real misollar bilan yozing; umumiy va bo'sh gaplardan saqlaning; matn mantiqiy va bir-biriga bog'liq bo'lsin; sodda, tushunarli va ta'lim standartlariga mos o'zbek tilida yozing."""


def _generate_with_groq(topic: str, wiki_context: str) -> Tuple[List[str], List[str]]:
    user_prompt = f"""Mavzu: {topic}
Qisqacha ma'lumot (tayanch sifatida, agar bo'sh bo'lsa o'z bilimingizdan foydalaning): {wiki_context}

Vazifa: Ushbu mavzu bo'yicha to'liq va professional taqdimot (prezentatsiya) tuzing.

Qat'iy talablar:
1) Aniq {MIN_SECTIONS} banddan iborat reja tuzing (raqamlangan 1-{MIN_SECTIONS}), oxirgi band albatta "Xulosa" bo'lsin. Bandlar mavzu turiga mos, mantiqiy ketma-ketlikda bo'lsin (tizim ko'rsatmasidagi struktura tamoyillariga qarang).
2) Har bir band uchun 250-400 so'zdan iborat, aniq va tushunarli matn yozing. Matnlar bir-biriga mos, mantiqiy ketma-ketlikda bo'lsin.
3) Matnlarda aniq faktlar, misollar va tushunchalar bo'lsin. Umumiy va bo'sh gaplardan saqlaning.
4) O'zbek tilida, sodda va ta'lim standartlariga mos uslubda yozing.

Natijani FAQAT quyidagi formatda qaytaring, boshqa hech qanday izoh yozmang:
Reja:
1. ...
2. ...
...
{MIN_SECTIONS}. ...
Matnlar:
1. ...
2. ...
...
{MIN_SECTIONS}. ...
"""
    try:
        resp = groq_client.chat.completions.create(
            model=GROQ_MODEL,
            messages=[
                {"role": "system", "content": GROQ_SYSTEM_PROMPT},
                {"role": "user", "content": user_prompt},
            ],
            temperature=0.5,
            max_tokens=8000,
        )
        raw = resp.choices[0].message.content.strip()
        raw = re.sub(r"[*#]", "", raw)
    except Exception:
        logging.exception("Groq AI xatosi:")
        return [], []

    plan, contents = _parse_plan_response(raw)
    if len(plan) < 4 or len(contents) < 4:
        return [], []
    while len(contents) < len(plan):
        contents.append("Ma'lumot yetishmadi.")
    return plan, contents


# ---------------- WIKIPEDIA CONTENT GENERATION (fallback, AI'siz) ----------------
SKIP_WIKI_SECTIONS = {
    "manbalar", "adabiyotlar", "havolalar", "izohlar", "tashqi havolalar",
    "shuningdek qarang", "yana qarang", "eslatmalar", "manba",
    "qo'shimcha adabiyotlar", "bibliografiya", "izoh",
}


def _clean_wiki_text(text: str) -> str:
    text = re.sub(r"\[\d+\]", "", text)  # citation markers like [1]
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{2,}", "\n", text)
    return text.strip()


def _first_sentences(text: str, n: int) -> str:
    sentences = re.split(r"(?<=[.!?])\s+", text.strip())
    return " ".join(sentences[:n]).strip()


def _fetch_wikipedia_content(topic: str) -> str:
    try:
        page = wikipedia.page(topic, auto_suggest=True)
        return page.content
    except wikipedia.DisambiguationError as e:
        try:
            page = wikipedia.page(e.options[0], auto_suggest=False)
            return page.content
        except Exception:
            logging.exception("Wikipedia disambiguation xatosi:")
            return ""
    except Exception:
        logging.exception("Wikipedia sahifasini olishda xatolik:")
        return ""


def _generate_from_wikipedia_only(topic: str) -> Tuple[List[str], List[str]]:
    """Groq ishlamay qolsa ishlatiladigan zaxira: reja va matnlarni Wikipedia maqolasidan (AI ishtirokisiz) tuzadi."""
    content = _fetch_wikipedia_content(topic)

    plan: List[str] = []
    contents: List[str] = []
    intro = ""
    sections: List[Tuple[str, str]] = []

    if content:
        parts = re.split(r"\n==+\s*(.+?)\s*==+\n", content)
        intro = _clean_wiki_text(parts[0])
        for i in range(1, len(parts) - 1, 2):
            heading = parts[i].strip()
            body = _clean_wiki_text(parts[i + 1])
            if heading.lower() in SKIP_WIKI_SECTIONS or len(body) < 50:
                continue
            sections.append((heading, body))

    if intro or sections:
        if intro:
            plan.append("Kirish")
            contents.append(intro)
        for heading, body in sections[:8]:
            plan.append(heading)
            contents.append(body)

        closing_source = intro or (sections[0][1] if sections else "")
        closing = _first_sentences(closing_source, 3) or f"{topic} mavzusi bo'yicha asosiy ma'lumotlar yuqorida keltirildi."
        plan.append("Xulosa")
        contents.append(closing)
    else:
        plan = ["Kirish", "Asosiy tushunchalar", "Amaliy misollar", "Muammolar va yechimlar", "Xulosa"]
        contents = [
            f"'{topic}' mavzusi bo'yicha Wikipedia'da maqola topilmadi. "
            "Iltimos, mavzuni aniqroq yoki boshqacha nom bilan qayta kiriting."
        ] * 5

    while len(contents) < len(plan):
        contents.append("Ma'lumot yetishmadi — iltimos mavzuni kengroq yozing.")

    return plan, contents


def generate_plan_and_contents(topic: str) -> Tuple[List[str], List[str]]:
    """Reja va matnlarni avvalo Groq AI orqali, muvaffaqiyatsiz bo'lsa Wikipedia'dan tuzadi."""
    try:
        wiki_context = wikipedia.summary(topic, sentences=8)
    except Exception:
        wiki_context = ""

    plan, contents = _generate_with_groq(topic, wiki_context)
    if plan and contents:
        return plan, contents

    return _generate_from_wikipedia_only(topic)


# ---------------- BOT HANDLERS ----------------
@dp.message(CommandStart())
async def cmd_start(message: types.Message, state: FSMContext):
    await state.clear()
    await message.answer(
        "👋 Salom! Men Prezentatsiya Botman.\n\n"
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


@dp.message(Form.waiting_for_image)
async def handle_skip_image(message: types.Message, state: FSMContext):
    text = (message.text or "").lower().strip()
    if text == "skip":
        data = await state.get_data()
        topic = data.get("topic")
        await state.clear()
        await message.reply("Prezentatsiya yaratilmoqda...")
        await create_presentation(message, topic, [], [])
    else:
        await message.reply("Iltimos, rasm yuboring yoki rasmsiz davom etish uchun 'skip' deb yozing.")


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
        content_chunks.extend(split_text_to_chunks(text, max_chars=1400))

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
