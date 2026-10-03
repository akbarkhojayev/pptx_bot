"""Taqdimot kontentini tayyorlash: avval Groq AI (tuzilgan JSON), ishlamasa -- Wikipedia (AI'siz)."""
import json
import logging
import os
import re
from typing import List, Optional, Tuple

import wikipedia
import wikipedia.wikipedia as _wikipedia_internal
from openai import OpenAI

from pptx_builder import DEFAULT_PALETTE, PALETTES

wikipedia.set_lang("uz")
# Wikimedia rate-limits the library's shared default User-Agent; a distinct one avoids collateral 429s.
_wikipedia_internal.USER_AGENT = "PptxPresentationBot/1.0 (Telegram content-generation bot)"

GROQ_API_KEY = os.environ.get("GROQ_API_KEY", "")
GROQ_MODEL = "openai/gpt-oss-120b"
MIN_CONTENT_SLIDES = 5

_groq_client = OpenAI(api_key=GROQ_API_KEY, base_url="https://api.groq.com/openai/v1") if GROQ_API_KEY else None

SLIDE_TYPES = ("bullets", "cards", "stats", "timeline", "compare", "quote", "conclusion")

# ---------------- PROMPTS ----------------
GROQ_SYSTEM_PROMPT = """Siz "PPT Yordamchi" — professional taqdimot dizayneri va kontent mutaxassisisiz.
Siz yozgan slaydlar haqiqiy konferensiya yoki dars taqdimotidagi kabi bo'ladi: har bir slaydda BITTA asosiy fikr,
qisqa va lo'nda punktlar (gap-devor emas), aniq faktlar, raqamlar va real misollar. Batafsil tushuntirish
slaydga emas, ma'ruzachi eslatmasiga (notes) yoziladi.

Avval mavzu turini aniqlang va strukturani unga moslang:
- Ilmiy/diplom: kirish, maqsad va vazifalar, nazariy asoslar, tahlil, natijalar, xulosa.
- Biznes/startup: muammo, yechim, bozor, ustunliklar, raqobatchilar bilan taqqoslash, rivojlanish rejasi, xulosa.
- Ta'lim/dars: kirish, asosiy tushunchalar, tarixi, turlari/tasnifi, amaliy qo'llanilishi, muammo va yechimlar, xulosa.

Faqat ishonchli, tekshirilgan faktlarni yozing. Raqam yoki sanani aniq bilmasangiz — uni o'ylab topmang,
boshqa slayd turini tanlang. Til: adabiy, sodda va tushunarli o'zbek tili (lotin yozuvi)."""

USER_PROMPT_TEMPLATE = """Mavzu: {topic}
Tayanch ma'lumot (bo'sh bo'lsa, o'z bilimingizdan foydalaning): {context}

Ushbu mavzu bo'yicha {min_slides}-{max_slides} ta kontent slayddan iborat taqdimot tuzing (titul, reja va
"rahmat" slaydlari avtomatik qo'shiladi, ularni yozmang).

Slayd turlari va maydonlari:
- "bullets": {{"type":"bullets","title":"...","bullets":["...", ...],"highlight":"...","notes":"..."}}
    3-5 ta punkt, har biri 60-110 belgi; "Atama: izoh" ko'rinishi ma'qul. "highlight" — slayddagi eng muhim
    bitta fikr yoki qiziqarli fakt (90-150 belgi), ixtiyoriy.
- "cards": {{"type":"cards","title":"...","items":[{{"title":"...","text":"..."}}, ...],"notes":"..."}}
    3-4 ta element (turlar, tamoyillar, afzalliklar). title ≤ 30 belgi, text 80-140 belgi.
- "stats": {{"type":"stats","title":"...","items":[{{"value":"...","label":"..."}}, ...],"text":"...","notes":"..."}}
    2-4 ta HAQIQIY raqam. value ≤ 8 belgi (masalan "1991", "8 mlrd", "45%"), label 30-80 belgi,
    text — raqamlar nimani anglatishi haqida 1-2 gap.
- "timeline": {{"type":"timeline","title":"...","items":[{{"label":"...","text":"..."}}, ...],"notes":"..."}}
    3-5 ta bosqich yoki sana, xronologik tartibda. label ≤ 12 belgi (yil yoki "1-bosqich"), text 50-110 belgi.
- "compare": {{"type":"compare","title":"...","left":{{"title":"...","points":[...]}},"right":{{"title":"...","points":[...]}},"notes":"..."}}
    Ikki tomonni taqqoslash (afzallik/kamchilik, oldin/keyin, A va B). Har tomonda 3-4 punkt, 40-90 belgi.
- "quote": {{"type":"quote","text":"...","author":"...","notes":"..."}}
    Faqat mavzuga oid HAQIQIY, mashhur iqtibos bo'lsa (≤ 160 belgi). Ishonchingiz komil bo'lmasa — ishlatmang.
- "conclusion": {{"type":"conclusion","title":"Xulosa","bullets":[...],"highlight":"...","notes":"..."}}
    Oxirgi slayd: 3-5 ta asosiy xulosa.

Qoidalar:
1) Birinchi slayd — "Kirish" (bullets), oxirgisi — "conclusion".
2) Kamida 4 xil turdan foydalaning, bir xil tur ketma-ket 2 martadan ortiq kelmasin.
3) Slayd sarlavhasi 25-60 belgi, mazmunli (masalan "Python qayerda ishlatiladi?"), oxirida nuqta yo'q.
4) "notes" — ma'ruzachi shu slaydda og'zaki aytadigan matn: 60-120 so'z, slayddagi punktlarni kengaytiradi.
5) Markdown belgilaridan (**, #, `) foydalanmang.
6) "palette" — mavzu ruhiga mos rang palitrasi: {palettes}.
   (masalan: tibbiyot/ekologiya — "teal" yoki "forest", tarix — "terracotta" yoki "berry",
   texnologiya — "ocean" yoki "midnight", biznes — "charcoal" yoki "midnight", san'at — "royal").

Natijani FAQAT bitta JSON obyekt sifatida qaytaring:
{{"title":"taqdimotning aniq nomi","subtitle":"bir gapli qisqa tavsif (≤ 110 belgi)","palette":"...","slides":[...]}}"""


# ---------------- CLEANING / VALIDATION ----------------
def _clean(text, limit: Optional[int] = None) -> str:
    if not isinstance(text, str):
        text = "" if text is None else str(text)
    text = re.sub(r"\*\*|__|`|^#+\s*", "", text, flags=re.M)
    text = re.sub(r"\s+", " ", text).strip()
    if limit and len(text) > limit:
        cut = text[:limit].rsplit(" ", 1)[0].rstrip(",;:-—")
        text = cut + "…"
    return text


def _clean_title(text, limit: int = 80) -> str:
    return _clean(text, limit).rstrip(".")


def _clean_list(items, limit: int, max_items: int) -> List[str]:
    if not isinstance(items, list):
        return []
    out = [_clean(x, limit) for x in items]
    return [x for x in out if x][:max_items]


def _normalize_slide(raw: dict) -> Optional[dict]:
    if not isinstance(raw, dict):
        return None
    typ = raw.get("type")
    title = _clean_title(raw.get("title"))
    notes = _clean(raw.get("notes"), 1500)
    s = {"type": typ, "title": title, "notes": notes}

    if typ in ("bullets", "conclusion"):
        s["bullets"] = _clean_list(raw.get("bullets"), 170, 6)
        s["highlight"] = _clean(raw.get("highlight"), 200)
        if len(s["bullets"]) < 2:
            return None
        if typ == "conclusion" and not title:
            s["title"] = "Xulosa"
    elif typ == "cards":
        items = []
        for it in (raw.get("items") or [])[:6]:
            if isinstance(it, dict) and _clean(it.get("title")) and _clean(it.get("text")):
                items.append({"title": _clean_title(it["title"], 45), "text": _clean(it["text"], 200)})
        if len(items) < 2:
            return None
        s["items"] = items
    elif typ == "stats":
        items = []
        for it in (raw.get("items") or [])[:4]:
            if isinstance(it, dict) and _clean(it.get("value")) and _clean(it.get("label")):
                items.append({"value": _clean(it["value"], 12), "label": _clean(it["label"], 110)})
        if len(items) < 2:
            return None
        s["items"] = items
        s["text"] = _clean(raw.get("text"), 300)
    elif typ == "timeline":
        items = []
        for it in (raw.get("items") or [])[:5]:
            if isinstance(it, dict) and _clean(it.get("label")) and _clean(it.get("text")):
                items.append({"label": _clean(it["label"], 16), "text": _clean(it["text"], 150)})
        if len(items) < 2:
            return None
        s["items"] = items
    elif typ == "compare":
        sides = []
        for key in ("left", "right"):
            side = raw.get(key)
            if not isinstance(side, dict):
                return None
            points = _clean_list(side.get("points"), 120, 5)
            if not points or not _clean(side.get("title")):
                return None
            sides.append({"title": _clean_title(side["title"], 40), "points": points})
        s["left"], s["right"] = sides
    elif typ == "quote":
        s["text"] = _clean(raw.get("text"), 220).strip("\"“”«»")
        s["author"] = _clean(raw.get("author"), 60)
        if not s["text"]:
            return None
        s["title"] = s["title"] or s["text"]
    else:
        return None

    if not s["title"]:
        return None
    return s


def normalize_deck(raw: dict, topic: str) -> Optional[dict]:
    if not isinstance(raw, dict):
        return None
    slides = [s for s in (_normalize_slide(x) for x in (raw.get("slides") or [])) if s]
    if len(slides) < MIN_CONTENT_SLIDES:
        return None
    palette = raw.get("palette") if raw.get("palette") in PALETTES else DEFAULT_PALETTE
    return {
        "title": _clean_title(raw.get("title"), 90) or topic,
        "subtitle": _clean(raw.get("subtitle"), 140),
        "palette": palette,
        "slides": slides[:14],
    }


def _extract_json(text: str) -> Optional[dict]:
    try:
        return json.loads(text)
    except Exception:
        pass
    start, end = text.find("{"), text.rfind("}")
    if start != -1 and end > start:
        try:
            return json.loads(text[start:end + 1])
        except Exception:
            return None
    return None


# ---------------- GROQ ----------------
def _generate_with_groq(topic: str, context: str) -> Optional[dict]:
    if _groq_client is None:
        logging.warning("GROQ_API_KEY o'rnatilmagan — Wikipedia zaxirasidan foydalaniladi.")
        return None
    user_prompt = USER_PROMPT_TEMPLATE.format(
        topic=topic, context=context or "—", min_slides=9, max_slides=12, palettes=", ".join(PALETTES),
    )
    messages = [
        {"role": "system", "content": GROQ_SYSTEM_PROMPT},
        {"role": "user", "content": user_prompt},
    ]
    # Avval JSON rejimida, u xato bersa (model formatni buzsa) oddiy rejimda qayta urinamiz.
    for kwargs in ({"response_format": {"type": "json_object"}}, {}):
        try:
            resp = _groq_client.chat.completions.create(
                model=GROQ_MODEL, messages=messages, temperature=0.5, max_tokens=12000, **kwargs,
            )
            raw = _extract_json(resp.choices[0].message.content or "")
            deck = normalize_deck(raw, topic)
            if deck:
                return deck
            logging.warning("Groq javobi yaroqsiz tuzilishga ega, qayta urinilmoqda.")
        except Exception:
            logging.exception("Groq AI xatosi:")
    return None


# ---------------- WIKIPEDIA FALLBACK ----------------
SKIP_WIKI_SECTIONS = {
    "manbalar", "adabiyotlar", "havolalar", "izohlar", "tashqi havolalar",
    "shuningdek qarang", "yana qarang", "eslatmalar", "manba",
    "qo'shimcha adabiyotlar", "bibliografiya", "izoh",
}
_YEAR_RE = re.compile(r"\b(1[0-9]{3}|20[0-9]{2})\b")


def _clean_wiki_text(text: str) -> str:
    text = re.sub(r"\[\d+\]", "", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{2,}", "\n", text)
    return text.strip()


def _sentences(text: str) -> List[str]:
    return [s.strip() for s in re.split(r"(?<=[.!?])\s+", text.replace("\n", " ")) if len(s.strip()) > 15]


def _fetch_wikipedia_content(topic: str) -> str:
    try:
        return wikipedia.page(topic, auto_suggest=True).content
    except wikipedia.DisambiguationError as e:
        try:
            return wikipedia.page(e.options[0], auto_suggest=False).content
        except Exception:
            logging.exception("Wikipedia disambiguation xatosi:")
    except Exception:
        logging.exception("Wikipedia sahifasini olishda xatolik:")
    return ""


def _section_slide(heading: str, body: str, allow_timeline: bool = True) -> Optional[dict]:
    sents = _sentences(body)
    if not sents:
        return None
    notes = _clean(body, 1200)
    dated = [(m.group(1), s) for s in sents if (m := _YEAR_RE.search(s))]
    if allow_timeline and len(dated) >= 4:
        items, seen = [], set()
        for year, s in dated:
            if year not in seen:
                seen.add(year)
                items.append({"label": year, "text": s})
        if len(items) >= 3:
            items.sort(key=lambda it: int(it["label"]))
            return _normalize_slide({"type": "timeline", "title": heading, "items": items[:5], "notes": notes})
    raw = {"type": "bullets", "title": heading, "bullets": sents[:4], "notes": notes}
    extra = [x for x in sents[4:] if 60 <= len(x) <= 200]
    if extra:
        raw["highlight"] = extra[0]
    return _normalize_slide(raw)


def _generate_from_wikipedia_only(topic: str) -> Optional[dict]:
    content = _fetch_wikipedia_content(topic)
    if not content:
        return None
    parts = re.split(r"\n==+\s*(.+?)\s*==+\n", content)
    intro = _clean_wiki_text(parts[0])
    slides = []
    intro_sents = _sentences(intro)
    if intro_sents:
        slides.append(_normalize_slide({"type": "bullets", "title": "Kirish", "bullets": intro_sents[:4],
                                        "notes": _clean(intro, 1200)}))
    for i in range(1, len(parts) - 1, 2):
        heading = parts[i].strip()
        body = _clean_wiki_text(parts[i + 1])
        if heading.lower() in SKIP_WIKI_SECTIONS or len(body) < 80:
            continue
        prev_timeline = bool(slides) and slides[-1] is not None and slides[-1]["type"] == "timeline"
        slides.append(_section_slide(heading, body, allow_timeline=not prev_timeline))
        if len(slides) >= 10:
            break
    slides = [s for s in slides if s]
    if not slides:
        return None
    closing = intro_sents[:3] or [f"{topic} mavzusi bo'yicha asosiy ma'lumotlar ko'rib chiqildi."]
    if len(closing) < 2:
        closing.append(f"{topic} mavzusini chuqurroq o'rganish tavsiya etiladi.")
    slides.append(_normalize_slide({"type": "conclusion", "title": "Xulosa", "bullets": closing,
                                    "notes": "Taqdimotdan asosiy xulosalar."}))
    return {"title": topic, "subtitle": "", "palette": DEFAULT_PALETTE, "slides": slides}


def generate_deck(topic: str) -> Tuple[Optional[dict], str]:
    """(deck, manba) qaytaradi; manba — "ai" yoki "wikipedia". Hech narsa topilmasa deck = None."""
    try:
        context = wikipedia.summary(topic, sentences=8)
    except Exception:
        context = ""
    deck = _generate_with_groq(topic, context)
    if deck:
        return deck, "ai"
    return _generate_from_wikipedia_only(topic), "wikipedia"
