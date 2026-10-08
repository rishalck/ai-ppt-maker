"""
AI POWERPOINT MAKER - SINGLE FILE
=================================

A single-file Gemini-powered PowerPoint generator.

Features
--------
- Gemini-powered presentation planning and writing
- JSON structured generation with graceful fallback
- Professional dark / light / corporate / neon / minimal themes
- Multiple slide layouts
- Pictures from Unsplash Source-style public image URLs
- Optional Gemini image generation hook
- Charts using matplotlib
- Tables, timelines, comparisons, quote slides, KPI cards
- Speaker notes where supported
- Automatic text fitting
- PPTX export
- Optional Streamlit web UI
- CLI mode
- No API keys hard-coded

Install
-------
pip install google-genai python-pptx pillow requests matplotlib python-dotenv streamlit

PowerShell
----------
$env:GEMINI_API_KEY="YOUR_KEY"

CLI:
    python ppt_maker.py

Web UI:
    streamlit run ppt_maker.py

Notes
-----
The image system intentionally uses public image URLs as a simple fallback.
For production, replace IMAGE_SEARCH_ENDPOINT / image_provider() with your
preferred licensed image provider. Do not use copyrighted images without
permission.
"""

from __future__ import annotations

import os
import io
import re
import sys
import json
import math
import textwrap
import hashlib
import random
import tempfile
import traceback
from pathlib import Path
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

import requests
from PIL import Image, ImageOps, ImageFilter, ImageEnhance

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.dml.color import RGBColor

try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass

try:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
except Exception:
    plt = None

try:
    from google import genai
    from google.genai import types
except Exception:
    genai = None
    types = None


# ============================================================================
# CONFIGURATION
# ============================================================================

APP_NAME = "AI PowerPoint Maker"
OUTPUT_DIR = Path("generated_ppts")
IMAGE_DIR = OUTPUT_DIR / "images"
CHART_DIR = OUTPUT_DIR / "charts"
OUTPUT_DIR.mkdir(exist_ok=True)
IMAGE_DIR.mkdir(exist_ok=True)
CHART_DIR.mkdir(exist_ok=True)

DEFAULT_MODEL = os.getenv("GEMINI_MODEL", "gemini-2.5-flash")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")

SLIDE_W = 13.333
SLIDE_H = 7.5

IMAGE_TIMEOUT = 20
MAX_IMAGE_BYTES = 8 * 1024 * 1024

FONT_HEAD = "Aptos Display"
FONT_BODY = "Aptos"

# Unsplash Source endpoint is convenient for demos. For production, use a
# licensed image provider and its official API.
IMAGE_SEARCH_ENDPOINT = "https://images.unsplash.com/photo-"

# ============================================================================
# COLORS
# ============================================================================

def rgb(hex_value: str) -> RGBColor:
    value = hex_value.strip().lstrip("#")
    return RGBColor(int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16))


THEMES = {
    "midnight": {
        "bg": "#08090D",
        "surface": "#11131A",
        "surface2": "#181B24",
        "text": "#F5F7FA",
        "muted": "#9BA3B4",
        "accent": "#7C5CFF",
        "accent2": "#29D3FF",
        "positive": "#36D399",
        "warning": "#FBBF24",
        "danger": "#FB7185",
        "line": "#292D3A",
    },
    "ocean": {
        "bg": "#06131C",
        "surface": "#0C202D",
        "surface2": "#123143",
        "text": "#F1FAFF",
        "muted": "#9AB5C5",
        "accent": "#36C5F0",
        "accent2": "#5271FF",
        "positive": "#2DD4BF",
        "warning": "#F5C451",
        "danger": "#FF708D",
        "line": "#24485B",
    },
    "forest": {
        "bg": "#07110C",
        "surface": "#0E1D14",
        "surface2": "#152A1E",
        "text": "#F1FFF5",
        "muted": "#9AB4A2",
        "accent": "#42E58A",
        "accent2": "#A6E65A",
        "positive": "#42E58A",
        "warning": "#F4C95D",
        "danger": "#FF7188",
        "line": "#264332",
    },
    "sunset": {
        "bg": "#160A0B",
        "surface": "#251112",
        "surface2": "#351719",
        "text": "#FFF7F3",
        "muted": "#C8A8A0",
        "accent": "#FF795F",
        "accent2": "#FFB454",
        "positive": "#6EE7B7",
        "warning": "#FFD166",
        "danger": "#FF5D73",
        "line": "#4B2527",
    },
    "light": {
        "bg": "#F7F8FA",
        "surface": "#FFFFFF",
        "surface2": "#EEF1F5",
        "text": "#111827",
        "muted": "#687386",
        "accent": "#5B4BDB",
        "accent2": "#0788C9",
        "positive": "#0C9A6A",
        "warning": "#B7791F",
        "danger": "#D6455D",
        "line": "#DCE1E8",
    },
    "corporate": {
        "bg": "#F4F7FB",
        "surface": "#FFFFFF",
        "surface2": "#EAF0F7",
        "text": "#172033",
        "muted": "#657089",
        "accent": "#1D4ED8",
        "accent2": "#0891B2",
        "positive": "#059669",
        "warning": "#D97706",
        "danger": "#DC2626",
        "line": "#D5DCE7",
    },
    "neon": {
        "bg": "#050507",
        "surface": "#0D0D12",
        "surface2": "#15151C",
        "text": "#FFFFFF",
        "muted": "#A1A1AA",
        "accent": "#D946EF",
        "accent2": "#22D3EE",
        "positive": "#A3E635",
        "warning": "#FACC15",
        "danger": "#FB7185",
        "line": "#292936",
    },
}

# ============================================================================
# DATA MODELS
# ============================================================================

@dataclass
class PresentationConfig:
    topic: str
    slides: int = 10
    audience: str = "college students"
    style: str = "premium modern"
    theme: str = "midnight"
    language: str = "English"
    include_images: bool = True
    include_charts: bool = True
    include_notes: bool = True
    image_provider: str = "unsplash"
    author: str = ""
    company: str = ""


@dataclass
class SlideData:
    type: str
    title: str = ""
    subtitle: str = ""
    bullets: List[str] = field(default_factory=list)
    visual: str = ""
    image_query: str = ""
    quote: str = ""
    author: str = ""
    left_title: str = ""
    right_title: str = ""
    left_items: List[str] = field(default_factory=list)
    right_items: List[str] = field(default_factory=list)
    stats: List[Dict[str, str]] = field(default_factory=list)
    timeline: List[Dict[str, str]] = field(default_factory=list)
    labels: List[str] = field(default_factory=list)
    values: List[float] = field(default_factory=list)
    notes: str = ""
    table_headers: List[str] = field(default_factory=list)
    table_rows: List[List[str]] = field(default_factory=list)


# ============================================================================
# UTILITIES
# ============================================================================

def clean_text(value: Any, default: str = "") -> str:
    if value is None:
        return default
    return str(value).strip()


def slugify(value: str) -> str:
    value = re.sub(r"[^a-zA-Z0-9\s_-]", "", value)
    value = re.sub(r"[\s_-]+", "_", value.strip())
    return value.lower()[:80] or "presentation"


def safe_float(value: Any, default: float = 0.0) -> float:
    try:
        return float(value)
    except Exception:
        return default


def clamp(value: float, low: float, high: float) -> float:
    return max(low, min(high, value))


def wrap_lines(text: str, width: int = 40) -> List[str]:
    return textwrap.wrap(
        clean_text(text),
        width=width,
        break_long_words=False,
        break_on_hyphens=False,
    ) or [""]


def hex_color(theme: Dict[str, str], key: str) -> RGBColor:
    return rgb(theme[key])


def ensure_font(run, size: float, color: RGBColor, bold: bool = False,
                font_name: str = FONT_BODY):
    run.font.name = font_name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color


def set_cell_text(cell, text, size=12, color=None, bold=False):
    tf = cell.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = clean_text(text)
    ensure_font(
        run,
        size,
        color or RGBColor(255, 255, 255),
        bold,
    )


# ============================================================================
# GEMINI CLIENT
# ============================================================================

class GeminiEngine:
    def __init__(self, api_key: Optional[str] = None, model: str = DEFAULT_MODEL):
        self.api_key = api_key or GEMINI_API_KEY
        self.model = model
        self.client = None

        if genai and self.api_key:
            self.client = genai.Client(api_key=self.api_key)

    @property
    def available(self) -> bool:
        return self.client is not None

    def generate_text(self, prompt: str, system: str = "") -> str:
        if not self.available:
            raise RuntimeError(
                "Gemini is not configured. Set GEMINI_API_KEY."
            )

        kwargs = {}
        if types:
            kwargs["config"] = types.GenerateContentConfig(
                system_instruction=system or None,
                temperature=0.7,
            )

        response = self.client.models.generate_content(
            model=self.model,
            contents=prompt,
            **kwargs,
        )
        return clean_text(getattr(response, "text", ""))

    def generate_json(self, prompt: str, system: str = "") -> Dict[str, Any]:
        if not self.available:
            raise RuntimeError(
                "Gemini is not configured. Set GEMINI_API_KEY."
            )

        config_kwargs = {
            "temperature": 0.65,
            "response_mime_type": "application/json",
        }

        if types:
            config = types.GenerateContentConfig(
                system_instruction=system or None,
                **config_kwargs,
            )
        else:
            config = None

        response = self.client.models.generate_content(
            model=self.model,
            contents=prompt,
            config=config,
        )

        raw = clean_text(getattr(response, "text", ""))

        raw = re.sub(r"^```json\s*", "", raw, flags=re.I)
        raw = re.sub(r"^```\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)

        try:
            return json.loads(raw)
        except Exception:
            match = re.search(r"\{.*\}", raw, flags=re.S)
            if match:
                return json.loads(match.group(0))
            raise ValueError("Gemini returned invalid JSON.")


# ============================================================================
# PRESENTATION PLANNER
# ============================================================================

PLANNER_SYSTEM = """
You are an elite presentation designer, information architect, and writer.

Create concise, human-quality PowerPoint content.

Rules:
1. Never write giant paragraphs.
2. Prefer 3-5 short bullets.
3. Every slide needs a clear purpose.
4. Vary slide types.
5. Use visuals whenever useful.
6. Do not fabricate statistics.
7. If a slide uses statistics, label them as illustrative unless the user
   supplied reliable figures.
8. Make the sequence tell a coherent story.
9. The first slide is always title.
10. The final slide is always conclusion.
11. Use professional language appropriate for the audience.
12. Keep titles short and strong.
13. Image queries should be concrete visual descriptions.
14. Speaker notes should expand the slide instead of repeating it.
"""


def default_slide(slide_type: str, title: str) -> SlideData:
    return SlideData(
        type=slide_type,
        title=title,
        subtitle="",
        bullets=[],
        visual="Clean visual supporting the main idea",
        image_query="professional abstract presentation background",
    )


def normalize_slide(raw: Dict[str, Any]) -> SlideData:
    data = SlideData(
        type=clean_text(raw.get("type"), "content").lower(),
        title=clean_text(raw.get("title")),
        subtitle=clean_text(raw.get("subtitle")),
        bullets=[clean_text(x) for x in raw.get("bullets", []) if clean_text(x)],
        visual=clean_text(raw.get("visual")),
        image_query=clean_text(raw.get("image_query")),
        quote=clean_text(raw.get("quote")),
        author=clean_text(raw.get("author")),
        left_title=clean_text(raw.get("left_title")),
        right_title=clean_text(raw.get("right_title")),
        left_items=[clean_text(x) for x in raw.get("left_items", [])],
        right_items=[clean_text(x) for x in raw.get("right_items", [])],
        stats=[],
        timeline=[],
        labels=[clean_text(x) for x in raw.get("labels", [])],
        values=[safe_float(x) for x in raw.get("values", [])],
        notes=clean_text(raw.get("notes")),
        table_headers=[clean_text(x) for x in raw.get("table_headers", [])],
        table_rows=[
            [clean_text(y) for y in row]
            for row in raw.get("table_rows", [])
            if isinstance(row, list)
        ],
    )

    for item in raw.get("stats", []):
        if isinstance(item, dict):
            data.stats.append({
                "number": clean_text(item.get("number")),
                "label": clean_text(item.get("label")),
            })

    for item in raw.get("timeline", []):
        if isinstance(item, dict):
            data.timeline.append({
                "year": clean_text(item.get("year")),
                "title": clean_text(item.get("title")),
                "description": clean_text(item.get("description")),
            })

    allowed = {
        "title", "content", "comparison", "stats", "timeline", "quote",
        "chart", "table", "section", "conclusion", "image",
    }

    if data.type not in allowed:
        data.type = "content"

    return data


def fallback_plan(config: PresentationConfig) -> List[SlideData]:
    n = max(5, config.slides)

    slides: List[SlideData] = []

    slides.append(SlideData(
        type="title",
        title=config.topic,
        subtitle=f"A practical presentation for {config.audience}",
        visual="Premium hero image related to the topic",
        image_query=config.topic,
        notes=f"Introduce {config.topic} and explain why it matters.",
    ))

    templates = [
        ("content", "Why this topic matters"),
        ("content", "The core idea"),
        ("comparison", "Two ways to think about it"),
        ("stats", "Key signals"),
        ("timeline", "How it evolved"),
        ("chart", "What the numbers suggest"),
        ("content", "Real-world applications"),
        ("image", "The bigger picture"),
        ("table", "Practical comparison"),
    ]

    for i in range(1, n - 1):
        t, title = templates[(i - 1) % len(templates)]
        s = default_slide(t, title)

        s.bullets = [
            f"Important aspect of {config.topic}",
            "A practical example makes the idea easier to understand",
            "The main takeaway is simple and actionable",
        ]
        s.image_query = config.topic
        s.notes = f"Explain {title.lower()} in the context of {config.topic}."

        if t == "comparison":
            s.left_title = "Approach A"
            s.right_title = "Approach B"
            s.left_items = ["Simple workflow", "Lower complexity", "Easy to start"]
            s.right_items = ["Advanced workflow", "More flexibility", "More setup"]

        if t == "stats":
            s.stats = [
                {"number": "01", "label": "Core idea"},
                {"number": "03", "label": "Key actions"},
                {"number": "05", "label": "Main outcomes"},
            ]

        if t == "timeline":
            s.timeline = [
                {"year": "01", "title": "Start", "description": "The problem appears."},
                {"year": "02", "title": "Change", "description": "New approaches emerge."},
                {"year": "03", "title": "Scale", "description": "Adoption grows."},
                {"year": "04", "title": "Next", "description": "New opportunities appear."},
            ]

        if t == "chart":
            s.labels = ["A", "B", "C", "D"]
            s.values = [25, 48, 67, 84]

        if t == "table":
            s.table_headers = ["Factor", "Option A", "Option B"]
            s.table_rows = [
                ["Complexity", "Low", "Medium"],
                ["Speed", "Fast", "Moderate"],
                ["Flexibility", "Good", "High"],
            ]

        slides.append(s)

    slides.append(SlideData(
        type="conclusion",
        title="Key takeaways",
        bullets=[
            f"{config.topic} is easier to understand when broken into clear ideas.",
            "Focus on the practical use rather than unnecessary complexity.",
            "Start small, measure results, and improve continuously.",
        ],
        visual="Minimal closing visual",
        image_query=config.topic,
        notes="Close with the three most important ideas and invite questions.",
    ))

    return slides


def build_plan(engine: GeminiEngine, config: PresentationConfig) -> List[SlideData]:
    prompt = f"""
Create a {config.slides}-slide PowerPoint presentation.

Topic: {config.topic}
Audience: {config.audience}
Style: {config.style}
Theme: {config.theme}
Language: {config.language}

Return exactly this JSON shape:

{{
  "slides": [
    {{
      "type": "title|content|comparison|stats|timeline|quote|chart|table|section|image|conclusion",
      "title": "short title",
      "subtitle": "optional short subtitle",
      "bullets": ["short point", "short point"],
      "visual": "visual description",
      "image_query": "specific image search phrase",
      "quote": "optional quote",
      "author": "optional author",
      "left_title": "optional",
      "left_items": ["..."],
      "right_title": "optional",
      "right_items": ["..."],
      "stats": [
        {{"number": "80%", "label": "label"}}
      ],
      "timeline": [
        {{"year": "2020", "title": "event", "description": "short"}}
      ],
      "labels": ["A", "B", "C"],
      "values": [20, 40, 60],
      "table_headers": ["A", "B"],
      "table_rows": [["x", "y"]],
      "notes": "speaker notes"
    }}
  ]
}}

Requirements:
- First slide type must be title.
- Last slide type must be conclusion.
- Mix layouts.
- Keep bullet text short.
- Do not invent factual statistics.
- For illustrative chart data, make the notes say "Illustrative data".
"""

    try:
        result = engine.generate_json(prompt, PLANNER_SYSTEM)
        raw_slides = result.get("slides", [])
        if not isinstance(raw_slides, list) or not raw_slides:
            return fallback_plan(config)

        slides = [normalize_slide(x) for x in raw_slides]
        slides = slides[:max(3, config.slides)]

        if slides[0].type != "title":
            slides.insert(0, SlideData(
                type="title",
                title=config.topic,
                subtitle="AI-generated presentation",
                image_query=config.topic,
            ))

        if slides[-1].type != "conclusion":
            slides.append(SlideData(
                type="conclusion",
                title="Key takeaways",
                bullets=[
                    "Understand the central idea",
                    "Focus on practical application",
                    "Keep learning and improving",
                ],
                notes="Summarize the presentation.",
            ))

        return slides

    except Exception:
        return fallback_plan(config)


# ============================================================================
# IMAGE ENGINE
# ============================================================================

# A small set of reliable public Unsplash photo IDs. These are only examples.
# Replace or expand with your licensed provider for production use.
IMAGE_LIBRARY = {
    "technology": [
        "photo-1518770660439-4636190af475",
        "photo-1550751827-4bd374c3f58b",
        "photo-1519389950473-47ba0277781c",
    ],
    "business": [
        "photo-1497366754035-f200968a6e72",
        "photo-1556761175-b413da4baf72",
        "photo-1521737711867-e3b97375f902",
    ],
    "education": [
        "photo-1523240795612-9a054b0db644",
        "photo-1503676260728-1c00da094a0b",
        "photo-1524178232363-1fb2b075b655",
    ],
    "ai": [
        "photo-1677442136019-21780ecad995",
        "photo-1620712943543-bcc4688e7485",
        "photo-1555255707-c07966088b7b",
    ],
    "health": [
        "photo-1576091160399-112ba8d25d1d",
        "photo-1532938911079-1b06ac7ceec7",
        "photo-1505751172876-fa1923c5c528",
    ],
    "nature": [
        "photo-1441974231531-c6227db76b6e",
        "photo-1500534623283-312aade485b7",
        "photo-1501854140801-50d01698950b",
    ],
    "finance": [
        "photo-1559526324-593bc073d938",
        "photo-1611974789855-9c2a0a7236a3",
        "photo-1526304640581-d334cdbbf45e",
    ],
    "default": [
        "photo-1517245386807-bb43f82c33c4",
        "photo-1497366811353-6870744d04b2",
        "photo-1516321318423-f06f85e504b3",
    ],
}


def choose_image_category(query: str) -> str:
    q = query.lower()

    categories = {
        "technology": ["technology", "software", "computer", "coding", "cloud", "cyber"],
        "business": ["business", "company", "startup", "marketing", "management"],
        "education": ["education", "student", "school", "college", "learning"],
        "ai": ["ai", "artificial intelligence", "machine learning", "robot", "neural"],
        "health": ["health", "medical", "doctor", "hospital", "fitness"],
        "nature": ["nature", "environment", "forest", "climate", "earth"],
        "finance": ["finance", "money", "bank", "investment", "stock"],
    }

    for category, words in categories.items():
        if any(word in q for word in words):
            return category

    return "default"


def image_url_for_query(query: str, width: int = 1600, height: int = 900) -> str:
    category = choose_image_category(query)
    photo_ids = IMAGE_LIBRARY.get(category, IMAGE_LIBRARY["default"])
    photo_id = photo_ids[
        int(hashlib.md5(query.encode("utf-8")).hexdigest(), 16) % len(photo_ids)
    ]

    return (
        f"https://images.unsplash.com/{photo_id}"
        f"?auto=format&fit=crop&w={width}&h={height}&q=85"
    )


def download_image(query: str, width=1600, height=900) -> Optional[Path]:
    if not query:
        query = "abstract professional presentation"

    filename = (
        slugify(query)[:45]
        + "_"
        + hashlib.md5(query.encode()).hexdigest()[:10]
        + ".jpg"
    )
    path = IMAGE_DIR / filename

    if path.exists():
        return path

    url = image_url_for_query(query, width, height)

    try:
        response = requests.get(
            url,
            timeout=IMAGE_TIMEOUT,
            headers={"User-Agent": "AI-PowerPoint-Maker/1.0"},
        )
        response.raise_for_status()

        if len(response.content) > MAX_IMAGE_BYTES:
            return None

        image = Image.open(io.BytesIO(response.content)).convert("RGB")
        image.thumbnail((width, height), Image.Resampling.LANCZOS)
        image.save(path, "JPEG", quality=88, optimize=True)
        return path

    except Exception:
        return None


def create_gradient_background(
    path: Path,
    width: int,
    height: int,
    color_a: Tuple[int, int, int],
    color_b: Tuple[int, int, int],
):
    image = Image.new("RGB", (width, height))
    px = image.load()

    for y in range(height):
        t = y / max(1, height - 1)
        r = int(color_a[0] * (1 - t) + color_b[0] * t)
        g = int(color_a[1] * (1 - t) + color_b[1] * t)
        b = int(color_a[2] * (1 - t) + color_b[2] * t)

        for x in range(width):
            px[x, y] = (r, g, b)

    image.save(path, "JPEG", quality=88)


def fallback_image(query: str, theme: Dict[str, str]) -> Path:
    name = "fallback_" + hashlib.md5(query.encode()).hexdigest()[:12] + ".jpg"
    path = IMAGE_DIR / name

    if path.exists():
        return path

    a = tuple(rgb(theme["accent"]))
    b = tuple(rgb(theme["accent2"]))

    create_gradient_background(path, 1400, 800, a, b)
    return path


def get_image(query: str, theme: Dict[str, str]) -> Path:
    path = download_image(query)
    if path:
        return path
    return fallback_image(query, theme)


# ============================================================================
# PPTX ENGINE
# ============================================================================

class PPTXBuilder:
    def __init__(self, config: PresentationConfig):
        self.config = config
        self.theme = THEMES.get(config.theme, THEMES["midnight"])

        self.prs = Presentation()
        self.prs.slide_width = Inches(SLIDE_W)
        self.prs.slide_height = Inches(SLIDE_H)

        self.slide_number = 0

    # ------------------------------------------------------------------------
    # Basic drawing
    # ------------------------------------------------------------------------

    def new_slide(self):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self.slide_number += 1
        self.background(slide)
        return slide

    def background(self, slide):
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = hex_color(self.theme, "bg")

    def add_shape(
        self,
        slide,
        shape_type,
        x,
        y,
        w,
        h,
        fill=None,
        line=None,
        radius=False,
    ):
        shape = slide.shapes.add_shape(
            shape_type,
            Inches(x),
            Inches(y),
            Inches(w),
            Inches(h),
        )

        if fill:
            shape.fill.solid()
            shape.fill.fore_color.rgb = hex_color(self.theme, fill) \
                if isinstance(fill, str) and fill in self.theme \
                else (fill if isinstance(fill, RGBColor) else rgb(fill))
        else:
            shape.fill.background()

        if line:
            shape.line.color.rgb = (
                hex_color(self.theme, line)
                if isinstance(line, str) and line in self.theme
                else rgb(line)
            )
        else:
            shape.line.fill.background()

        return shape

    def rounded_box(self, slide, x, y, w, h, fill="surface", line=None):
        return self.add_shape(
            slide,
            MSO_SHAPE.ROUNDED_RECTANGLE,
            x,
            y,
            w,
            h,
            fill=fill,
            line=line,
        )

    def line(self, slide, x1, y1, x2, y2, color="line", width=1.2):
        shape = slide.shapes.add_connector(
            1,
            Inches(x1),
            Inches(y1),
            Inches(x2),
            Inches(y2),
        )
        shape.line.color.rgb = hex_color(self.theme, color)
        shape.line.width = Pt(width)
        return shape

    def text(
        self,
        slide,
        value,
        x,
        y,
        w,
        h,
        size=18,
        color="text",
        bold=False,
        align=PP_ALIGN.LEFT,
        valign=MSO_ANCHOR.TOP,
        font=FONT_BODY,
        margin=0.04,
    ):
        box = slide.shapes.add_textbox(
            Inches(x),
            Inches(y),
            Inches(w),
            Inches(h),
        )

        tf = box.text_frame
        tf.clear()
        tf.word_wrap = True
        tf.vertical_anchor = valign
        tf.margin_left = Inches(margin)
        tf.margin_right = Inches(margin)
        tf.margin_top = Inches(margin)
        tf.margin_bottom = Inches(margin)

        p = tf.paragraphs[0]
        p.alignment = align

        run = p.add_run()
        run.text = clean_text(value)

        ensure_font(
            run,
            size,
            hex_color(self.theme, color) if color in self.theme else rgb(color),
            bold,
            font,
        )

        return box

    def rich_text(
        self,
        slide,
        segments,
        x,
        y,
        w,
        h,
        size=18,
        align=PP_ALIGN.LEFT,
    ):
        box = slide.shapes.add_textbox(
            Inches(x),
            Inches(y),
            Inches(w),
            Inches(h),
        )

        tf = box.text_frame
        tf.clear()
        tf.word_wrap = True

        p = tf.paragraphs[0]
        p.alignment = align

        for segment in segments:
            run = p.add_run()
            run.text = segment.get("text", "")
            ensure_font(
                run,
                segment.get("size", size),
                hex_color(self.theme, segment.get("color", "text")),
                segment.get("bold", False),
                segment.get("font", FONT_BODY),
            )

        return box

    def add_picture_cover(self, slide, path, x, y, w, h, opacity=False):
        try:
            image = Image.open(path).convert("RGB")
            target_w = int(w * 144)
            target_h = int(h * 144)
            fitted = ImageOps.fit(
                image,
                (target_w, target_h),
                method=Image.Resampling.LANCZOS,
            )

            temp = IMAGE_DIR / (
                "fit_"
                + hashlib.md5(
                    f"{path}{x}{y}{w}{h}".encode()
                ).hexdigest()[:16]
                + ".jpg"
            )

            fitted.save(temp, "JPEG", quality=90)
            slide.shapes.add_picture(
                str(temp),
                Inches(x),
                Inches(y),
                width=Inches(w),
                height=Inches(h),
            )

        except Exception:
            return False

        return True

    def image_card(self, slide, query, x, y, w, h):
        path = get_image(query, self.theme)
        self.add_picture_cover(slide, path, x, y, w, h)

        overlay = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(x),
            Inches(y + h * 0.72),
            Inches(w),
            Inches(h * 0.28),
        )
        overlay.fill.solid()
        overlay.fill.fore_color.rgb = hex_color(self.theme, "bg")
        overlay.fill.transparency = 22
        overlay.line.fill.background()

    def footer(self, slide):
        self.text(
            slide,
            f"{self.slide_number:02d}",
            12.25,
            7.05,
            0.45,
            0.22,
            size=9,
            color="muted",
            align=PP_ALIGN.RIGHT,
        )

    def accent_bar(self, slide, x=0.72, y=0.62, h=0.58):
        self.add_shape(
            slide,
            MSO_SHAPE.ROUNDED_RECTANGLE,
            x,
            y,
            0.07,
            h,
            fill="accent",
        )

    def title(self, slide, title, subtitle=""):
        self.accent_bar(slide)

        self.text(
            slide,
            title,
            0.95,
            0.52,
            10.8,
            0.68,
            size=30,
            bold=True,
            font=FONT_HEAD,
        )

        if subtitle:
            self.text(
                slide,
                subtitle,
                0.96,
                1.23,
                10.8,
                0.42,
                size=12,
                color="muted",
            )

    # ------------------------------------------------------------------------
    # Slide layouts
    # ------------------------------------------------------------------------

    def slide_title(self, data: SlideData):
        slide = self.new_slide()

        if self.config.include_images:
            path = get_image(data.image_query or self.config.topic, self.theme)
            self.add_picture_cover(slide, path, 7.15, 0, 6.183, SLIDE_H)

            overlay = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Inches(6.7),
                0,
                Inches(6.633),
                Inches(SLIDE_H),
            )
            overlay.fill.solid()
            overlay.fill.fore_color.rgb = hex_color(self.theme, "bg")
            overlay.fill.transparency = 25
            overlay.line.fill.background()

        self.text(
            slide,
            self.config.company or "AI PRESENTATION",
            0.85,
            0.75,
            4.5,
            0.35,
            size=10,
            color="accent",
            bold=True,
        )

        self.text(
            slide,
            data.title or self.config.topic,
            0.85,
            1.75,
            6.0,
            1.65,
            size=42,
            bold=True,
            font=FONT_HEAD,
        )

        if data.subtitle:
            self.text(
                slide,
                data.subtitle,
                0.9,
                3.65,
                5.65,
                0.9,
                size=18,
                color="muted",
            )

        self.add_shape(
            slide,
            MSO_SHAPE.ROUNDED_RECTANGLE,
            0.9,
            5.2,
            1.1,
            0.08,
            fill="accent",
        )

        if self.config.author:
            self.text(
                slide,
                self.config.author,
                0.9,
                5.65,
                5.0,
                0.35,
                size=11,
                color="muted",
            )

        self.footer(slide)
        self.add_notes(slide, data.notes)
        return slide

    def slide_content(self, data: SlideData):
        slide = self.new_slide()
        self.title(slide, data.title, data.subtitle)

        bullets = data.bullets[:5]

        y = 2.0
        for i, bullet in enumerate(bullets):
            self.rounded_box(
                slide,
                0.82,
                y,
                7.0,
                0.72,
                "surface",
            )

            self.text(
                slide,
                f"{i + 1:02d}",
                1.05,
                y + 0.18,
                0.42,
                0.25,
                size=10,
                color="accent",
                bold=True,
            )

            self.text(
                slide,
                bullet,
                1.58,
                y + 0.12,
                5.95,
                0.45,
                size=15,
            )

            y += 0.87

        if self.config.include_images:
            self.image_card(
                slide,
                data.image_query or self.config.topic,
                8.25,
                1.95,
                4.2,
                4.35,
            )
        else:
            self.rounded_box(
                slide,
                8.25,
                1.95,
                4.2,
                4.35,
                "surface",
            )
            self.text(
                slide,
                data.visual or "Visual",
                8.65,
                3.25,
                3.4,
                1.5,
                size=20,
                color="muted",
                align=PP_ALIGN.CENTER,
                valign=MSO_ANCHOR.MIDDLE,
            )

        self.footer(slide)
        self.add_notes(slide, data.notes)
        return slide

    def slide_comparison(self, data: SlideData):
        slide = self.new_slide()
        self.title(slide, data.title, data.subtitle)

        self.rounded_box(slide, 0.75, 1.95, 5.8, 4.65, "surface")
        self.rounded_box(slide, 6.78, 1.95, 5.8, 4.65, "surface")

        self.text(
            slide,
            data.left_title or "Option A",
            1.15,
            2.3,
            4.8,
            0.5,
            size=22,
            bold=True,
            color="accent2",
        )

        self.text(
            slide,
            data.right_title or "Option B",
            7.18,
            2.3,
            4.8,
            0.5,
            size=22,
            bold=True,
            color="accent",
        )

        for idx, item in enumerate(data.left_items[:6]):
            self.text(
                slide,
                "• " + item,
                1.15,
                3.05 + idx * 0.53,
                4.9,
                0.42,
                size=14,
            )

        for idx, item in enumerate(data.right_items[:6]):
            self.text(
                slide,
                "• " + item,
                7.18,
                3.05 + idx * 0.53,
                4.9,
                0.42,
                size=14,
            )

        self.footer(slide)
        self.add_notes(slide, data.notes)
        return slide

    def slide_stats(self, data: SlideData):
        slide = self.new_slide()
        self.title(slide, data.title, data.subtitle)

        stats = data.stats[:4]

        if not stats:
            stats = [
                {"number": "01", "label": "Core idea"},
                {"number": "03", "label": "Key actions"},
                {"number": "05", "label": "Main outcomes"},
            ]

        gap = 0.25
        card_w = (11.8 - gap * (len(stats) - 1)) / max(1, len(stats))

        for i, stat in enumerate(stats):
            x = 0.75 + i * (card_w + gap)

            self.rounded_box(
                slide,
                x,
                2.05,
                card_w,
                3.25,
                "surface",
            )

            self.text(
                slide,
                stat.get("number", ""),
                x + 0.32,
                2.55,
                card_w - 0.64,
                0.9,
                size=34,
                bold=True,
                color="accent",
                font=FONT_HEAD,
            )

            self.text(
                slide,
                stat.get("label", ""),
                x + 0.32,
                3.7,
                card_w - 0.64,
                0.9,
                size=14,
                color="muted",
            )

        self.footer(slide)
        self.add_notes(slide, data.notes)
        return slide

    def slide_timeline(self, data: SlideData):
        slide = self.new_slide()
        self.title(slide, data.title, data.subtitle)

        items = data.timeline[:5]

        if not items:
            items = [
                {"year": "01", "title": "Start", "description": "The idea begins."},
                {"year": "02", "title": "Growth", "description": "Adoption increases."},
                {"year": "03", "title": "Scale", "description": "The system matures."},
                {"year": "04", "title": "Next", "description": "New opportunities emerge."},
            ]

        y_line = 4.1
        self.line(slide, 1.1, y_line, 12.15, y_line, "line", 2.5)

        width = 10.9 / max(1, len(items))

        for i, item in enumerate(items):
            x = 0.95 + i * width

            dot = self.add_shape(
                slide,
                MSO_SHAPE.OVAL,
                x + width / 2 - 0.1,
                y_line - 0.1,
                0.2,
                0.2,
                fill="accent",
            )

            self.text(
                slide,
                item.get("year", ""),
                x,
                2.25,
                width - 0.15,
                0.35,
                size=12,
                color="accent",
                bold=True,
                align=PP_ALIGN.CENTER,
            )

            self.text(
                slide,
                item.get("title", ""),
                x,
                2.85,
                width - 0.15,
                0.65,
                size=17,
                bold=True,
                align=PP_ALIGN.CENTER,
            )

            self.text(
                slide,
                item.get("description", ""),
                x + 0.05,
                4.55,
                width - 0.25,
                1.1,
                size=11,
                color="muted",
                align=PP_ALIGN.CENTER,
            )

        self.footer(slide)
        self.add_notes(slide, data.notes)
        return slide

    def slide_quote(self, data: SlideData):
        slide = self.new_slide()
        self.title(slide, data.title)

        self.text(
            slide,
            "“",
            1.05,
            1.65,
            1.0,
            0.8,
            size=62,
            color="accent",
            bold=True,
            font=FONT_HEAD,
        )

        self.text(
            slide,
            data.quote or "A strong idea becomes powerful when people can understand it.",
            1.55,
            2.35,
            9.8,
            2.0,
            size=30,
            bold=True,
            font=FONT_HEAD,
        )

        self.text(
            slide,
            "— " + (data.author or "Perspective"),
            1.6,
            4.9,
            7.0,
            0.45,
            size=14,
            color="muted",
        )

        self.footer(slide)
        self.add_notes(slide, data.notes)
        return slide

    def slide_chart(self, data: SlideData):
        slide = self.new_slide()
        self.title(slide, data.title, data.subtitle)

        chart_path = self.make_chart(data)

        if chart_path:
            slide.shapes.add_picture(
                str(chart_path),
                Inches(0.85),
                Inches(1.75),
                width=Inches(8.1),
                height=Inches(4.9),
            )

        self.rounded_box(slide, 9.25, 2.0, 3.25, 3.9, "surface")

        self.text(
            slide,
            "READ THE CHART",
            9.62,
            2.4,
            2.4,
            0.35,
            size=10,
            color="accent",
            bold=True,
        )

        self.text(
            slide,
            data.visual or "Illustrative visualization generated from the slide plan.",
            9.62,
            3.0,
            2.45,
            1.9,
            size=15,
            color="muted",
        )

        self.footer(slide)
        self.add_notes(slide, data.notes)
        return slide

    def make_chart(self, data: SlideData) -> Optional[Path]:
        if plt is None:
            return None

        labels = data.labels or ["A", "B", "C", "D"]
        values = data.values or [25, 40, 58, 75]

        values = values[:8]
        labels = labels[:len(values)]

        path = CHART_DIR / (
            slugify(data.title or "chart")
            + "_"
            + hashlib.md5(
                json.dumps(values).encode()
            ).hexdigest()[:8]
            + ".png"
        )

        if path.exists():
            return path

        fig = plt.figure(figsize=(10, 5.5), dpi=160)
        ax = fig.add_subplot(111)

        accent = self.theme["accent"]
        accent_hex = accent

        ax.bar(labels, values, color=accent_hex)
        ax.set_facecolor(self.theme["surface"])
        fig.patch.set_facecolor(self.theme["bg"])

        ax.tick_params(colors=self.theme["muted"], labelsize=10)
        for spine in ax.spines.values():
            spine.set_visible(False)

        ax.grid(
            axis="y",
            alpha=0.15,
            color=self.theme["muted"],
        )
        ax.set_axisbelow(True)

        fig.tight_layout()
        fig.savefig(path, facecolor=fig.get_facecolor(), bbox_inches="tight")
        plt.close(fig)

        return path

    def slide_table(self, data: SlideData):
        slide = self.new_slide()
        self.title(slide, data.title, data.subtitle)

        headers = data.table_headers or ["Factor", "Option A", "Option B"]
        rows = data.table_rows or [
            ["Speed", "Fast", "Moderate"],
            ["Complexity", "Low", "Medium"],
            ["Flexibility", "Good", "High"],
        ]

        rows = rows[:7]
        cols = len(headers)

        table_shape = slide.shapes.add_table(
            len(rows) + 1,
            cols,
            Inches(0.8),
            Inches(1.95),
            Inches(11.75),
            Inches(4.5),
        )

        table = table_shape.table

        for c, header in enumerate(headers):
            cell = table.cell(0, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = hex_color(self.theme, "accent")
            set_cell_text(
                cell,
                header,
                12,
                RGBColor(255, 255, 255),
                True,
            )

        for r, row in enumerate(rows, start=1):
            for c in range(cols):
                cell = table.cell(r, c)
                cell.fill.solid()

                if r % 2:
                    cell.fill.fore_color.rgb = hex_color(
                        self.theme,
                        "surface",
                    )
                else:
                    cell.fill.fore_color.rgb = hex_color(
                        self.theme,
                        "surface2",
                    )

                value = row[c] if c < len(row) else ""
                set_cell_text(
                    cell,
                    value,
                    11,
                    hex_color(self.theme, "text"),
                    False,
                )

        self.footer(slide)
        self.add_notes(slide, data.notes)
        return slide

    def slide_section(self, data: SlideData):
        slide = self.new_slide()

        self.text(
            slide,
            "SECTION",
            0.9,
            1.3,
            2.0,
            0.4,
            size=11,
            color="accent",
            bold=True,
        )

        self.text(
            slide,
            data.title,
            0.9,
            2.0,
            10.5,
            1.25,
            size=40,
            bold=True,
            font=FONT_HEAD,
        )

        self.text(
            slide,
            data.subtitle or data.visual,
            0.95,
            3.55,
            8.8,
            1.0,
            size=17,
            color="muted",
        )

        if self.config.include_images:
            self.image_card(
                slide,
                data.image_query or self.config.topic,
                9.5,
                1.1,
                2.8,
                4.8,
            )

        self.footer(slide)
        self.add_notes(slide, data.notes)
        return slide

    def slide_image(self, data: SlideData):
        slide = self.new_slide()

        if self.config.include_images:
            path = get_image(data.image_query or self.config.topic, self.theme)
            self.add_picture_cover(
                slide,
                path,
                0,
                0,
                SLIDE_W,
                SLIDE_H,
            )

            overlay = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                0,
                Inches(4.85),
                Inches(SLIDE_W),
                Inches(2.65),
            )
            overlay.fill.solid()
            overlay.fill.fore_color.rgb = hex_color(self.theme, "bg")
            overlay.fill.transparency = 12
            overlay.line.fill.background()

        self.text(
            slide,
            data.title,
            0.85,
            5.25,
            10.8,
            0.75,
            size=32,
            bold=True,
            font=FONT_HEAD,
        )

        self.text(
            slide,
            data.subtitle or data.visual,
            0.9,
            6.05,
            10.2,
            0.55,
            size=13,
            color="muted",
        )

        self.footer(slide)
        self.add_notes(slide, data.notes)
        return slide

    def slide_conclusion(self, data: SlideData):
        slide = self.new_slide()

        self.text(
            slide,
            "TAKEAWAYS",
            0.9,
            1.0,
            3.0,
            0.35,
            size=11,
            color="accent",
            bold=True,
        )

        self.text(
            slide,
            data.title or "Key takeaways",
            0.9,
            1.65,
            10.2,
            0.95,
            size=40,
            bold=True,
            font=FONT_HEAD,
        )

        bullets = data.bullets[:4]

        for i, bullet in enumerate(bullets):
            y = 3.0 + i * 0.7

            self.text(
                slide,
                "→",
                1.0,
                y,
                0.45,
                0.35,
                size=17,
                color="accent",
                bold=True,
            )

            self.text(
                slide,
                bullet,
                1.55,
                y,
                9.5,
                0.45,
                size=17,
            )

        self.text(
            slide,
            "Thank you",
            0.95,
            6.45,
            3.0,
            0.4,
            size=12,
            color="muted",
        )

        self.footer(slide)
        self.add_notes(slide, data.notes)
        return slide

    # ------------------------------------------------------------------------
    # Notes
    # ------------------------------------------------------------------------

    def add_notes(self, slide, notes: str):
        if not self.config.include_notes or not notes:
            return

        try:
            notes_slide = slide.notes_slide
            text_frame = notes_slide.notes_text_frame
            text_frame.text = notes
        except Exception:
            pass

    # ------------------------------------------------------------------------
    # Build
    # ------------------------------------------------------------------------

    def render_slide(self, data: SlideData):
        mapping = {
            "title": self.slide_title,
            "content": self.slide_content,
            "comparison": self.slide_comparison,
            "stats": self.slide_stats,
            "timeline": self.slide_timeline,
            "quote": self.slide_quote,
            "chart": self.slide_chart,
            "table": self.slide_table,
            "section": self.slide_section,
            "image": self.slide_image,
            "conclusion": self.slide_conclusion,
        }

        renderer = mapping.get(data.type, self.slide_content)

        try:
            return renderer(data)
        except Exception:
            # A malformed slide should not kill the entire presentation.
            fallback = SlideData(
                type="content",
                title=data.title or "Slide",
                bullets=data.bullets or ["Content unavailable"],
                visual=data.visual,
                image_query=data.image_query,
                notes=data.notes,
            )
            return self.slide_content(fallback)

    def build(self, slides: List[SlideData]) -> Path:
        for slide in slides:
            self.render_slide(slide)

        filename = slugify(self.config.topic) + ".pptx"
        output = OUTPUT_DIR / filename
        self.prs.save(output)

        return output


# ============================================================================
# AI IMAGE / CONTENT ENHANCEMENT
# ============================================================================

def improve_image_queries(engine: GeminiEngine, slides: List[SlideData]) -> List[SlideData]:
    if not engine.available:
        return slides

    payload = [
        {
            "title": s.title,
            "visual": s.visual,
            "image_query": s.image_query,
        }
        for s in slides
        if s.type not in ("title", "conclusion")
    ]

    prompt = f"""
Improve image search queries for these presentation slides.

Return JSON:
{{"items":[{{"title":"slide title","image_query":"specific photographic query"}}]}}

Rules:
- Prefer realistic photography, editorial visuals, architecture, people,
  objects, environments, or clean abstract scenes.
- Do not include text in the requested image.
- Do not request logos.
- Keep each query under 12 words.

Slides:
{json.dumps(payload, ensure_ascii=False)}
"""

    try:
        result = engine.generate_json(prompt, PLANNER_SYSTEM)
        items = result.get("items", [])

        by_title = {
            clean_text(x.get("title")).lower(): clean_text(x.get("image_query"))
            for x in items
            if isinstance(x, dict)
        }

        for slide in slides:
            q = by_title.get(slide.title.lower())
            if q:
                slide.image_query = q

    except Exception:
        pass

    return slides


def improve_slide_copy(engine: GeminiEngine, slides: List[SlideData]) -> List[SlideData]:
    if not engine.available:
        return slides

    compact = [
        {
            "title": s.title,
            "bullets": s.bullets,
            "subtitle": s.subtitle,
        }
        for s in slides
        if s.type in ("content", "conclusion")
    ]

    prompt = f"""
Polish this PowerPoint copy.

Return JSON:
{{"slides":[{{"title":"short title","subtitle":"short subtitle",
"bullets":["short bullet"]}}]}}

Rules:
- Preserve meaning.
- No marketing fluff.
- No paragraph bullets.
- Maximum 5 bullets per slide.
- Maximum roughly 14 words per bullet.

Slides:
{json.dumps(compact, ensure_ascii=False)}
"""

    try:
        result = engine.generate_json(prompt, PLANNER_SYSTEM)
        items = result.get("slides", [])

        index = 0
        for slide in slides:
            if slide.type in ("content", "conclusion") and index < len(items):
                item = items[index]
                slide.title = clean_text(item.get("title"), slide.title)
                slide.subtitle = clean_text(item.get("subtitle"), slide.subtitle)
                slide.bullets = [
                    clean_text(x)
                    for x in item.get("bullets", slide.bullets)
                    if clean_text(x)
                ][:5]
                index += 1

    except Exception:
        pass

    return slides


# ============================================================================
# MAIN GENERATION PIPELINE
# ============================================================================

def generate_presentation(config: PresentationConfig) -> Tuple[Path, List[SlideData]]:
    engine = GeminiEngine()

    slides = build_plan(engine, config)

    if config.include_images:
        slides = improve_image_queries(engine, slides)

    slides = improve_slide_copy(engine, slides)

    builder = PPTXBuilder(config)
    output = builder.build(slides)

    return output, slides


# ============================================================================
# CLI
# ============================================================================

def ask(prompt: str, default: str = "") -> str:
    suffix = f" [{default}]" if default else ""
    value = input(prompt + suffix + ": ").strip()
    return value or default


def ask_int(prompt: str, default: int) -> int:
    value = ask(prompt, str(default))
    try:
        return max(3, min(50, int(value)))
    except Exception:
        return default


def cli():
    print()
    print("=" * 65)
    print("             AI POWERPOINT MAKER")
    print("=" * 65)
    print()

    if not GEMINI_API_KEY:
        print("WARNING: GEMINI_API_KEY is not set.")
        print('PowerShell: $env:GEMINI_API_KEY="YOUR_KEY"')
        print()

    topic = ask("Presentation topic")
    if not topic:
        print("Topic is required.")
        return

    slides = ask_int("Number of slides", 10)
    audience = ask("Audience", "college students")
    style = ask("Style", "premium modern")
    theme = ask(
        "Theme: midnight/ocean/forest/sunset/light/corporate/neon",
        "midnight",
    )

    if theme not in THEMES:
        theme = "midnight"

    language = ask("Language", "English")
    author = ask("Author", "")
    company = ask("Company / college", "")

    config = PresentationConfig(
        topic=topic,
        slides=slides,
        audience=audience,
        style=style,
        theme=theme,
        language=language,
        author=author,
        company=company,
        include_images=True,
        include_charts=True,
        include_notes=True,
    )

    print()
    print("Generating presentation...")
    print("1. Planning slides")
    print("2. Improving content")
    print("3. Preparing pictures")
    print("4. Building PowerPoint")
    print()

    try:
        output, generated_slides = generate_presentation(config)

        print("=" * 65)
        print("DONE")
        print("=" * 65)
        print()
        print(f"Slides: {len(generated_slides)}")
        print(f"File:   {output.resolve()}")
        print()

    except Exception as exc:
        print()
        print("Generation failed:")
        print(str(exc))
        print()
        traceback.print_exc()


# ============================================================================
# STREAMLIT UI
# ============================================================================

def streamlit_app():
    import streamlit as st

    st.set_page_config(
        page_title="AI PPT Maker — AI Presentation Studio",
        page_icon="✦",
        layout="wide",
        initial_sidebar_state="collapsed",
    )

    # Inject ultra-premium styling
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap');

        /* Reset & Root Variables */
        :root {
            --bg-base: #07080C;
            --bg-surface: #0E1118;
            --bg-surface-elevated: #141822;
            --border-subtle: rgba(255, 255, 255, 0.08);
            --border-highlight: rgba(99, 102, 241, 0.35);
            --text-primary: #F8FAFC;
            --text-secondary: #94A3B8;
            --text-muted: #64748B;
            --accent-indigo: #6366F1;
            --accent-cyan: #06B6D4;
            --accent-violet: #8B5CF6;
        }

        /* Hide default Streamlit fluff */
        #MainMenu, footer, header { visibility: hidden; }
        .stDeployButton { display: none; }
        div[data-testid="stDecoration"] { display: none; }
        div[data-testid="stToolbar"] { display: none; }
        
        /* App Container & Background */
        .stApp {
            background-color: var(--bg-base);
            background-image: 
                radial-gradient(circle 900px at 50% -120px, rgba(99, 102, 241, 0.12), transparent),
                radial-gradient(circle 600px at 85% 300px, rgba(6, 182, 212, 0.05), transparent),
                radial-gradient(circle 600px at 15% 400px, rgba(139, 92, 246, 0.05), transparent);
            color: var(--text-primary);
            font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
            letter-spacing: -0.01em;
        }

        .block-container {
            padding-top: 1.25rem !important;
            padding-bottom: 5rem !important;
            max-width: 1240px !important;
        }

        /* Navigation Bar */
        .nav-container {
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 14px 24px;
            background: rgba(14, 17, 24, 0.7);
            backdrop-filter: blur(20px);
            -webkit-backdrop-filter: blur(20px);
            border: 1px solid var(--border-subtle);
            border-radius: 16px;
            margin-bottom: 48px;
            box-shadow: 0 8px 32px rgba(0, 0, 0, 0.35);
        }

        .nav-brand {
            display: flex;
            align-items: center;
            gap: 12px;
            text-decoration: none;
        }

        .nav-logo-icon {
            width: 32px;
            height: 32px;
            background: linear-gradient(135deg, #6366F1 0%, #8B5CF6 100%);
            border-radius: 8px;
            display: flex;
            align-items: center;
            justify-content: center;
            color: #ffffff;
            font-size: 16px;
            font-weight: 700;
            box-shadow: 0 0 16px rgba(99, 102, 241, 0.4);
        }

        .nav-title {
            font-size: 17px;
            font-weight: 700;
            letter-spacing: -0.02em;
            color: #FFFFFF;
        }

        .nav-badge {
            font-size: 11px;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            padding: 4px 10px;
            border-radius: 100px;
            background: rgba(99, 102, 241, 0.12);
            color: #A5B4FC;
            border: 1px solid rgba(99, 102, 241, 0.25);
            display: inline-flex;
            align-items: center;
            gap: 6px;
        }

        .nav-badge-dot {
            width: 6px;
            height: 6px;
            border-radius: 50%;
            background: #34D399;
            box-shadow: 0 0 8px #34D399;
        }

        .nav-links {
            display: flex;
            align-items: center;
            gap: 20px;
        }

        .nav-link {
            font-size: 14px;
            font-weight: 500;
            color: var(--text-secondary);
            text-decoration: none;
            transition: color 0.15s ease;
        }

        .nav-link:hover {
            color: var(--text-primary);
        }

        .nav-btn {
            font-size: 13px;
            font-weight: 600;
            color: #FFFFFF;
            background: rgba(255, 255, 255, 0.07);
            border: 1px solid rgba(255, 255, 255, 0.12);
            padding: 7px 14px;
            border-radius: 8px;
            text-decoration: none;
            transition: all 0.2s ease;
        }

        .nav-btn:hover {
            background: rgba(255, 255, 255, 0.12);
            border-color: rgba(255, 255, 255, 0.22);
            color: #FFFFFF;
        }

        /* Hero Header Section */
        .hero-header {
            text-align: center;
            max-width: 820px;
            margin: 0 auto 36px auto;
        }

        .hero-eyebrow {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            font-size: 12px;
            font-weight: 700;
            letter-spacing: 0.12em;
            text-transform: uppercase;
            color: #A5B4FC;
            background: rgba(99, 102, 241, 0.08);
            border: 1px solid rgba(99, 102, 241, 0.2);
            padding: 6px 14px;
            border-radius: 100px;
            margin-bottom: 22px;
        }

        .hero-headline {
            font-size: 52px;
            line-height: 1.12;
            font-weight: 800;
            letter-spacing: -0.035em;
            color: #FFFFFF;
            margin: 0 0 18px 0;
        }

        .hero-headline .highlight {
            background: linear-gradient(135deg, #FFFFFF 30%, #CBD5E1 70%, #94A3B8 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            text-shadow: 0 0 40px rgba(99, 102, 241, 0.25);
            font-style: normal;
        }

        .hero-subtext {
            font-size: 18px;
            line-height: 1.55;
            color: var(--text-secondary);
            font-weight: 400;
            max-width: 620px;
            margin: 0 auto;
        }

        /* Creation Box Container */
        .creation-container {
            background: linear-gradient(180deg, rgba(17, 21, 30, 0.85) 0%, rgba(12, 15, 22, 0.95) 100%);
            border: 1px solid var(--border-subtle);
            border-radius: 20px;
            padding: 24px;
            box-shadow: 0 20px 50px -10px rgba(0, 0, 0, 0.6), inset 0 1px 0 rgba(255, 255, 255, 0.06);
            backdrop-filter: blur(24px);
            -webkit-backdrop-filter: blur(24px);
            margin-bottom: 40px;
            position: relative;
        }

        .creation-container::before {
            content: '';
            position: absolute;
            top: 0;
            left: 20%;
            right: 20%;
            height: 1px;
            background: linear-gradient(90deg, transparent, rgba(99, 102, 241, 0.5), transparent);
        }

        /* Preset Chips */
        .chips-label {
            font-size: 12px;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            color: var(--text-muted);
            margin-bottom: 8px;
        }

        /* Streamlit Text Area Styling */
        div[data-testid="stTextArea"] textarea {
            background: rgba(8, 10, 15, 0.8) !important;
            border: 1px solid rgba(255, 255, 255, 0.1) !important;
            border-radius: 14px !important;
            color: #FFFFFF !important;
            font-size: 16px !important;
            line-height: 1.5 !important;
            font-family: 'Plus Jakarta Sans', sans-serif !important;
            padding: 16px 18px !important;
            box-shadow: inset 0 2px 4px rgba(0, 0, 0, 0.4) !important;
            transition: all 0.2s ease !important;
        }

        div[data-testid="stTextArea"] textarea:focus {
            border-color: rgba(99, 102, 241, 0.6) !important;
            box-shadow: 0 0 0 3px rgba(99, 102, 241, 0.15), inset 0 2px 4px rgba(0, 0, 0, 0.4) !important;
        }

        div[data-testid="stTextArea"] label {
            color: var(--text-primary) !important;
            font-size: 14px !important;
            font-weight: 600 !important;
            margin-bottom: 6px !important;
        }

        /* Streamlit Input / Select / Slider Controls */
        div[data-testid="stSelectbox"] label,
        div[data-testid="stSlider"] label,
        div[data-testid="stTextInput"] label {
            color: var(--text-secondary) !important;
            font-size: 12px !important;
            font-weight: 600 !important;
            text-transform: uppercase !important;
            letter-spacing: 0.04em !important;
        }

        div[data-testid="stSelectbox"] > div > div {
            background: rgba(8, 10, 15, 0.8) !important;
            border: 1px solid rgba(255, 255, 255, 0.08) !important;
            border-radius: 10px !important;
            color: #FFFFFF !important;
            font-size: 14px !important;
        }

        div[data-testid="stTextInput"] input {
            background: rgba(8, 10, 15, 0.8) !important;
            border: 1px solid rgba(255, 255, 255, 0.08) !important;
            border-radius: 10px !important;
            color: #FFFFFF !important;
            font-size: 14px !important;
            padding: 8px 12px !important;
        }

        /* Checkbox styling */
        div[data-testid="stCheckbox"] label {
            color: var(--text-secondary) !important;
            font-size: 13px !important;
            font-weight: 500 !important;
        }

        /* Primary Generate Button */
        div.stButton > button[kind="primary"] {
            background: linear-gradient(135deg, #4F46E5 0%, #6366F1 50%, #8B5CF6 100%) !important;
            color: #FFFFFF !important;
            font-size: 16px !important;
            font-weight: 700 !important;
            letter-spacing: -0.01em !important;
            border: 1px solid rgba(255, 255, 255, 0.2) !important;
            border-radius: 14px !important;
            padding: 16px 28px !important;
            box-shadow: 0 4px 20px rgba(99, 102, 241, 0.4), inset 0 1px 0 rgba(255, 255, 255, 0.25) !important;
            transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1) !important;
            cursor: pointer !important;
        }

        div.stButton > button[kind="primary"]:hover {
            transform: translateY(-2px) !important;
            box-shadow: 0 8px 30px rgba(99, 102, 241, 0.6), inset 0 1px 0 rgba(255, 255, 255, 0.35) !important;
            border-color: rgba(255, 255, 255, 0.35) !important;
        }

        div.stButton > button[kind="primary"]:active {
            transform: translateY(0px) !important;
        }

        /* Secondary & Download Buttons */
        div.stDownloadButton > button {
            background: linear-gradient(135deg, #10B981 0%, #059669 100%) !important;
            color: #FFFFFF !important;
            font-size: 16px !important;
            font-weight: 700 !important;
            border: 1px solid rgba(255, 255, 255, 0.2) !important;
            border-radius: 12px !important;
            padding: 14px 24px !important;
            box-shadow: 0 4px 20px rgba(16, 185, 129, 0.35) !important;
            transition: all 0.2s ease !important;
        }

        div.stDownloadButton > button:hover {
            transform: translateY(-1px) !important;
            box-shadow: 0 6px 24px rgba(16, 185, 129, 0.5) !important;
        }

        /* Expander */
        .streamlit-expanderHeader {
            background: rgba(14, 17, 24, 0.6) !important;
            border: 1px solid var(--border-subtle) !important;
            border-radius: 10px !important;
            color: var(--text-primary) !important;
            font-weight: 600 !important;
            font-size: 14px !important;
        }

        /* Slide Preview Cards (3D Fan & Visuals) */
        .preview-showcase {
            background: rgba(11, 14, 20, 0.6);
            border: 1px solid var(--border-subtle);
            border-radius: 20px;
            padding: 24px;
            margin-bottom: 48px;
            position: relative;
            overflow: hidden;
        }

        .preview-header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 20px;
        }

        .preview-tag {
            font-size: 11px;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.08em;
            color: #38BDF8;
            background: rgba(56, 189, 248, 0.1);
            border: 1px solid rgba(56, 189, 248, 0.2);
            padding: 4px 10px;
            border-radius: 6px;
        }

        .slide-deck-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
            gap: 16px;
        }

        .slide-mockup {
            background: #08090D;
            border: 1px solid rgba(255, 255, 255, 0.1);
            border-radius: 12px;
            padding: 16px;
            aspect-ratio: 16 / 9;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.5);
            transition: all 0.25s ease;
            position: relative;
            overflow: hidden;
        }

        .slide-mockup:hover {
            transform: translateY(-4px);
            border-color: rgba(99, 102, 241, 0.4);
            box-shadow: 0 16px 35px -8px rgba(99, 102, 241, 0.25);
        }

        .slide-mockup-title {
            font-size: 13px;
            font-weight: 700;
            color: #FFFFFF;
            margin-bottom: 4px;
            line-height: 1.25;
        }

        .slide-mockup-body {
            font-size: 10px;
            color: #94A3B8;
            line-height: 1.4;
        }

        .slide-mockup-badge {
            font-size: 9px;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            color: #818CF8;
            background: rgba(99, 102, 241, 0.15);
            padding: 2px 6px;
            border-radius: 4px;
            align-self: flex-start;
            margin-bottom: 8px;
        }

        .slide-mockup-kpi {
            display: flex;
            gap: 8px;
            margin-top: 8px;
        }

        .kpi-chip {
            background: rgba(255, 255, 255, 0.04);
            border: 1px solid rgba(255, 255, 255, 0.08);
            border-radius: 6px;
            padding: 4px 8px;
            flex: 1;
        }

        .kpi-chip-val {
            font-size: 12px;
            font-weight: 800;
            color: #38BDF8;
        }

        .kpi-chip-lbl {
            font-size: 8px;
            color: #64748B;
        }

        /* Section Story Cards */
        .section-header-wrap {
            text-align: center;
            margin: 64px auto 32px auto;
            max-width: 640px;
        }

        .section-eyebrow {
            font-size: 11px;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.1em;
            color: #818CF8;
            margin-bottom: 8px;
        }

        .section-title {
            font-size: 32px;
            font-weight: 800;
            letter-spacing: -0.025em;
            color: #FFFFFF;
            margin: 0 0 10px 0;
        }

        .section-desc {
            font-size: 15px;
            color: var(--text-secondary);
            margin: 0;
        }

        /* Workflow Step Grid */
        .workflow-grid {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 16px;
            margin-bottom: 48px;
        }

        .workflow-card {
            background: rgba(14, 17, 24, 0.7);
            border: 1px solid var(--border-subtle);
            border-radius: 14px;
            padding: 20px;
            transition: all 0.2s ease;
        }

        .workflow-card:hover {
            border-color: rgba(99, 102, 241, 0.3);
            transform: translateY(-2px);
        }

        .workflow-num {
            font-family: 'JetBrains Mono', monospace;
            font-size: 12px;
            font-weight: 600;
            color: #6366F1;
            margin-bottom: 10px;
        }

        .workflow-name {
            font-size: 16px;
            font-weight: 700;
            color: #FFFFFF;
            margin-bottom: 6px;
        }

        .workflow-text {
            font-size: 13px;
            color: var(--text-secondary);
            line-height: 1.45;
            margin: 0;
        }

        /* Features Grid */
        .feature-grid {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 20px;
            margin-bottom: 56px;
        }

        .feature-card {
            background: linear-gradient(180deg, rgba(17, 21, 30, 0.7) 0%, rgba(11, 14, 20, 0.85) 100%);
            border: 1px solid var(--border-subtle);
            border-radius: 16px;
            padding: 24px;
            transition: all 0.2s ease;
        }

        .feature-card:hover {
            border-color: rgba(99, 102, 241, 0.35);
        }

        .feature-icon-box {
            width: 38px;
            height: 38px;
            border-radius: 10px;
            background: rgba(99, 102, 241, 0.12);
            border: 1px solid rgba(99, 102, 241, 0.25);
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 18px;
            margin-bottom: 16px;
        }

        .feature-heading {
            font-size: 16px;
            font-weight: 700;
            color: #FFFFFF;
            margin-bottom: 8px;
        }

        .feature-body {
            font-size: 13px;
            line-height: 1.5;
            color: var(--text-secondary);
            margin: 0;
        }

        /* Status & Generation Live Feedback */
        .generation-live-card {
            background: rgba(14, 18, 26, 0.9);
            border: 1px solid rgba(99, 102, 241, 0.3);
            border-radius: 16px;
            padding: 20px 24px;
            margin: 20px 0;
            box-shadow: 0 10px 30px rgba(99, 102, 241, 0.15);
        }

        .status-step-row {
            display: flex;
            align-items: center;
            gap: 12px;
            margin: 8px 0;
            font-size: 14px;
            color: var(--text-secondary);
        }

        .status-step-active {
            color: #A5B4FC;
            font-weight: 600;
        }

        .status-step-done {
            color: #34D399;
        }

        /* Footer */
        .footer-wrap {
            border-top: 1px solid var(--border-subtle);
            padding: 32px 0 16px 0;
            margin-top: 64px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            color: var(--text-muted);
            font-size: 13px;
        }

        /* Responsive Breakpoints */
        @media (max-width: 900px) {
            .hero-headline { font-size: 38px; }
            .workflow-grid { grid-template-columns: repeat(2, 1fr); }
            .feature-grid { grid-template-columns: 1fr; }
        }

        @media (max-width: 600px) {
            .hero-headline { font-size: 30px; }
            .workflow-grid { grid-template-columns: 1fr; }
            .nav-container { flex-direction: column; gap: 12px; }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    # Top Navigation Bar
    st.markdown(
        """
        <div class="nav-container">
            <div class="nav-brand">
                <div class="nav-logo-icon">✦</div>
                <div class="nav-title">AI PPT Maker</div>
                <div class="nav-badge"><span class="nav-badge-dot"></span> Studio 2.5</div>
            </div>
            <div class="nav-links">
                <a href="#creation-studio" class="nav-link">Composer</a>
                <a href="#features" class="nav-link">Capabilities</a>
                <a href="#workflow" class="nav-link">Pipeline</a>
                <a href="https://github.com/rishalck/ai-ppt-maker" target="_blank" class="nav-btn">GitHub ↗</a>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Hero Header
    st.markdown(
        """
        <div class="hero-header">
            <div class="hero-eyebrow">
                <span>✦</span> AI PRESENTATION STUDIO
            </div>
            <h1 class="hero-headline">
                Turn an idea into a <br><span class="highlight">presentation</span> worth presenting.
            </h1>
            <p class="hero-subtext">
                Describe your idea. Our AI builds the story, designs the slides, and exports the deck.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Initialize prompt presets in session state if needed
    if "prompt_input" not in st.session_state:
        st.session_state.prompt_input = ""

    # Quick inspiration prompt chips
    st.markdown('<div class="chips-label">Inspiration Presets</div>', unsafe_allow_html=True)
    chip_col1, chip_col2, chip_col3, chip_col4 = st.columns(4)
    with chip_col1:
        if st.button("🧬 AI in Healthcare", use_container_width=True):
            st.session_state.prompt_input = "Artificial Intelligence in Healthcare: Breakthroughs in Clinical Diagnostics and Drug Discovery"
            st.rerun()
    with chip_col2:
        if st.button("🚀 Series A Pitch Deck", use_container_width=True):
            st.session_state.prompt_input = "Next-Gen Developer Platform: Seed-to-Series A Pitch Deck covering Problem, Traction, and Market Opportunity"
            st.rerun()
    with chip_col3:
        if st.button("⚡ Renewable Energy 2030", use_container_width=True):
            st.session_state.prompt_input = "Global Renewable Energy Transition: Solar, Wind, and Next-Gen Grid Infrastructure by 2030"
            st.rerun()
    with chip_col4:
        if st.button("📊 Q3 Growth Strategy", use_container_width=True):
            st.session_state.prompt_input = "Q3 Enterprise SaaS Go-To-Market Strategy, Revenue Scaling, and Customer Retention"
            st.rerun()

    # The Creation Box (Centerpiece)
    st.markdown('<div id="creation-studio" class="creation-container">', unsafe_allow_html=True)
    
    topic = st.text_area(
        "What do you want to present?",
        value=st.session_state.prompt_input,
        placeholder="e.g., Create a 10-slide presentation about the future of autonomous vehicles, focusing on safety milestones, commercial fleets, and urban infrastructure...",
        height=110,
        label_visibility="visible",
    )

    # Compact Control Strip
    c1, c2, c3, c4 = st.columns([1.2, 1.4, 1.4, 1.2])
    with c1:
        slides = st.slider("Slides", min_value=3, max_value=30, value=10)
    with c2:
        audience = st.selectbox(
            "Audience",
            [
                "Business executives",
                "College students",
                "Startup investors",
                "Technical professionals",
                "School students",
                "General audience",
            ],
            index=0,
        )
    with c3:
        style = st.selectbox(
            "Style",
            [
                "Premium modern",
                "Editorial minimal",
                "Corporate bold",
                "Startup pitch",
                "Creative narrative",
                "Academic rigor",
                "Futuristic dark",
            ],
            index=0,
        )
    with c4:
        theme = st.selectbox(
            "Theme",
            list(THEMES.keys()),
            index=0,
        )

    # Secondary Compact Options
    opt_col1, opt_col2, opt_col3, opt_col4 = st.columns([1.2, 1.1, 1.1, 1.2])
    with opt_col1:
        language = st.selectbox(
            "Language",
            ["English", "Malayalam", "Hindi", "Tamil", "Kannada", "Spanish", "German", "French"],
            index=0,
        )
    with opt_col2:
        include_images = st.checkbox("Include Pictures", value=True)
    with opt_col3:
        include_charts = st.checkbox("Include Charts", value=True)
    with opt_col4:
        include_notes = st.checkbox("Speaker Notes", value=True)

    # Expandable Presenter Credentials
    with st.expander("Optional Presenter Metadata (Author / Organization)"):
        meta_c1, meta_c2 = st.columns(2)
        with meta_c1:
            author = st.text_input("Author Name", placeholder="e.g. Alex Morgan")
        with meta_c2:
            company = st.text_input("Organization / Institution", placeholder="e.g. Stanford University / TechCorp")

    # Primary Generation Action Button
    st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
    generate = st.button("✦ Generate Presentation", type="primary", use_container_width=True)

    st.markdown('</div>', unsafe_allow_html=True)

    # Generation Workflow Execution
    if generate:
        if not topic.strip():
            st.error("Please enter a presentation topic or idea.")
            return

        config = PresentationConfig(
            topic=topic,
            slides=slides,
            audience=audience,
            style=style,
            theme=theme,
            language=language,
            include_images=include_images,
            include_charts=include_charts,
            include_notes=include_notes,
            author=author,
            company=company,
        )

        status_container = st.container()

        with status_container:
            st.markdown(
                """
                <div class="generation-live-card">
                    <div style="font-weight: 700; font-size: 16px; margin-bottom: 12px; color: #FFFFFF; display: flex; align-items: center; gap: 8px;">
                        <span>✦</span> Generating Presentation Studio Deck
                    </div>
                """,
                unsafe_allow_html=True,
            )
            prog_bar = st.progress(10)
            status_text = st.empty()

        try:
            status_text.markdown('<div class="status-step-row status-step-active"><span>①</span> Building your story and narrative architecture...</div>', unsafe_allow_html=True)
            prog_bar.progress(20)

            engine = GeminiEngine()
            generated_slides = build_plan(engine, config)

            prog_bar.progress(45)
            status_text.markdown('<div class="status-step-row status-step-active"><span>②</span> Refining slide copy and executive summaries...</div>', unsafe_allow_html=True)
            generated_slides = improve_slide_copy(engine, generated_slides)

            prog_bar.progress(65)
            if include_images:
                status_text.markdown('<div class="status-step-row status-step-active"><span>③</span> Curating high-resolution visual photography...</div>', unsafe_allow_html=True)
                generated_slides = improve_image_queries(engine, generated_slides)

            prog_bar.progress(85)
            status_text.markdown('<div class="status-step-row status-step-active"><span>④</span> Compiling native PowerPoint (.pptx) deck...</div>', unsafe_allow_html=True)

            builder = PPTXBuilder(config)
            output = builder.build(generated_slides)

            prog_bar.progress(100)
            status_text.markdown(f'<div class="status-step-row status-step-done"><span>✓</span> Complete! Successfully built {len(generated_slides)} presentation slides.</div>', unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

            # Presentation Ready Action Banner
            st.markdown("<div style='height: 16px;'></div>", unsafe_allow_html=True)
            st.success(f"✦ Presentation Generated: **{len(generated_slides)} slides** styled in **{theme.title()}** theme.")

            with open(output, "rb") as file:
                pptx_data = file.read()
                st.download_button(
                    label=f"⬇ Download Presentation ({output.name})",
                    data=pptx_data,
                    file_name=output.name,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True,
                )

            # Generated Deck Inspector
            st.markdown("<div style='height: 24px;'></div>", unsafe_allow_html=True)
            st.markdown("### Generated Slide System")
            
            for index, slide in enumerate(generated_slides, start=1):
                badge_type = slide.type.upper()
                with st.expander(f"Slide {index:02d} — {slide.title or 'Untitled Slide'}  [{badge_type}]", expanded=(index == 1)):
                    if slide.subtitle:
                        st.markdown(f"**Subtitle:** *{slide.subtitle}*")
                    
                    if slide.bullets:
                        st.markdown("**Key Takeaways:**")
                        for b in slide.bullets:
                            st.markdown(f"- {b}")

                    if slide.visual:
                        st.markdown(f"**Visual Concept:** `{slide.visual}`")

                    if slide.notes:
                        st.markdown(f"**Speaker Notes:** *\"{slide.notes}\"*")

        except Exception as exc:
            st.error(f"Generation error: {str(exc)}")
            st.code(traceback.format_exc())

    # Hero Visual / Live Slide Preview Fan Showcase
    st.markdown(
        """
        <div class="preview-showcase">
            <div class="preview-header">
                <div>
                    <span class="preview-tag">Studio Showcase</span>
                    <h3 style="font-size: 18px; font-weight: 700; color: #FFFFFF; margin: 6px 0 2px 0;">Realistic Slide Engine</h3>
                </div>
                <div style="font-size: 12px; color: #94A3B8;">16:9 Widescreen • Vector Typography</div>
            </div>
            
            <div class="slide-deck-grid">
                <!-- Slide 1: Keynote Title -->
                <div class="slide-mockup">
                    <div>
                        <div class="slide-mockup-badge">Keynote Title</div>
                        <div class="slide-mockup-title">Autonomous Mobility 2030</div>
                        <div class="slide-mockup-body">A comprehensive briefing on neural architectures and urban infrastructure.</div>
                    </div>
                    <div style="font-size: 9px; color: #64748B; border-top: 1px solid rgba(255,255,255,0.08); padding-top: 6px;">
                        Presented by Alex Morgan • Stanford AI Lab
                    </div>
                </div>

                <!-- Slide 2: Narrative & Imagery -->
                <div class="slide-mockup">
                    <div>
                        <div class="slide-mockup-badge">Visual Narrative</div>
                        <div class="slide-mockup-title">Perception & Sensor Fusion</div>
                        <div class="slide-mockup-body">• Sub-millisecond LiDAR clustering<br>• Real-time scene segmentation</div>
                    </div>
                    <div style="font-size: 9px; color: #38BDF8; font-weight: 600;">
                        ✦ High-Res Curated Visual
                    </div>
                </div>

                <!-- Slide 3: KPI Metrics -->
                <div class="slide-mockup">
                    <div>
                        <div class="slide-mockup-badge">Executive Metrics</div>
                        <div class="slide-mockup-title">Operational Validation</div>
                        <div class="slide-mockup-kpi">
                            <div class="kpi-chip"><div class="kpi-chip-val">99.4%</div><div class="kpi-chip-lbl">Accuracy</div></div>
                            <div class="kpi-chip"><div class="kpi-chip-val">+142%</div><div class="kpi-chip-lbl">Speed</div></div>
                            <div class="kpi-chip"><div class="kpi-chip-val">$4.2M</div><div class="kpi-chip-lbl">Efficiency</div></div>
                        </div>
                    </div>
                    <div style="font-size: 9px; color: #64748B;">Validated on 1.2M simulation miles</div>
                </div>

                <!-- Slide 4: Strategic Comparison -->
                <div class="slide-mockup">
                    <div>
                        <div class="slide-mockup-badge">Strategic Matrix</div>
                        <div class="slide-mockup-title">Approach Comparison</div>
                        <div class="slide-mockup-body">
                            <span style="color:#34D399;">✓ Approach A:</span> Low latency, edge compute<br>
                            <span style="color:#FBBF24;">⚡ Approach B:</span> High throughput, cloud scale
                        </div>
                    </div>
                    <div style="font-size: 9px; color: #64748B;">Multi-criteria decision framework</div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Section 1: "From idea to deck" (AI Workflow Pipeline)
    st.markdown(
        """
        <div id="workflow" class="section-header-wrap">
            <div class="section-eyebrow">The Generation Pipeline</div>
            <h2 class="section-title">From idea to deck in seconds</h2>
            <p class="section-desc">Four synchronized AI phases engineered to produce presentation-ready decks.</p>
        </div>

        <div class="workflow-grid">
            <div class="workflow-card">
                <div class="workflow-num">PHASE 01</div>
                <div class="workflow-name">Idea Ingestion</div>
                <p class="workflow-text">Gemini analyzes your topic, target audience, and intent to establish a strategic thesis.</p>
            </div>
            <div class="workflow-card">
                <div class="workflow-num">PHASE 02</div>
                <div class="workflow-name">Story Arc Architecture</div>
                <p class="workflow-text">Generates structured slide sequences with varied layouts: timelines, comparisons, and KPIs.</p>
            </div>
            <div class="workflow-card">
                <div class="workflow-num">PHASE 03</div>
                <div class="workflow-name">Visual & Data Synthesis</div>
                <p class="workflow-text">Curates photography and renders theme-harmonized Matplotlib charts and tables.</p>
            </div>
            <div class="workflow-card">
                <div class="workflow-num">PHASE 04</div>
                <div class="workflow-name">PowerPoint Compilation</div>
                <p class="workflow-text">Exports native .pptx files with vector shapes, exact font hierarchies, and speaker notes.</p>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Section 2: "Built for better slides" (Capabilities Grid)
    st.markdown(
        """
        <div id="features" class="section-header-wrap">
            <div class="section-eyebrow">Studio Capabilities</div>
            <h2 class="section-title">Built for better slides</h2>
            <p class="section-desc">Designed to replace generic templates with bespoke, intelligent presentations.</p>
        </div>

        <div class="feature-grid">
            <div class="feature-card">
                <div class="feature-icon-box">🧠</div>
                <div class="feature-heading">AI Story Structure</div>
                <p class="feature-body">Engaging narrative progression crafted specifically for executives, clients, or students without fluffy filler text.</p>
            </div>
            <div class="feature-card">
                <div class="feature-icon-box">📐</div>
                <div class="feature-heading">Intelligent Layouts</div>
                <p class="feature-body">Adaptive text fitting, multi-column comparisons, quote highlights, and structured timelines that look hand-crafted.</p>
            </div>
            <div class="feature-card">
                <div class="feature-icon-box">📊</div>
                <div class="feature-heading">Data & Charts</div>
                <p class="feature-body">Automatically synthesizes figures and generates clean, modern charts styled to your deck's color palette.</p>
            </div>
            <div class="feature-card">
                <div class="feature-icon-box">🖼</div>
                <div class="feature-heading">Curated Visuals</div>
                <p class="feature-body">Context-aware image selection matched with professional aspect ratios and aesthetic treatment.</p>
            </div>
            <div class="feature-card">
                <div class="feature-icon-box">🎙</div>
                <div class="feature-heading">Executive Speaker Notes</div>
                <p class="feature-body">Every slide includes tailored talking points and contextual cues embedded directly into the PPTX notes panel.</p>
            </div>
            <div class="feature-card">
                <div class="feature-icon-box">⚡</div>
                <div class="feature-heading">Universal PPTX Export</div>
                <p class="feature-body">100% editable Microsoft PowerPoint files compatible with Keynote, Google Slides, and PowerPoint 365.</p>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Section 3: "Your presentation, ready to present"
    st.markdown(
        """
        <div class="preview-showcase" style="text-align: center; padding: 48px 24px; background: linear-gradient(180deg, rgba(14,18,26,0.9) 0%, rgba(8,10,15,0.95) 100%);">
            <div class="section-eyebrow">Ready for the Boardroom</div>
            <h2 style="font-size: 30px; font-weight: 800; color: #FFFFFF; margin: 10px 0 14px 0;">Your presentation, ready to present.</h2>
            <p style="font-size: 15px; color: #94A3B8; max-width: 580px; margin: 0 auto 28px auto;">
                Export native widescreen decks with pixel-perfect contrast, legible typography, and zero watermark lock-in.
            </p>
            <div style="display: inline-flex; gap: 16px; flex-wrap: wrap; justify-content: center;">
                <span style="font-size: 12px; color: #CBD5E1; background: rgba(255,255,255,0.06); padding: 8px 16px; border-radius: 8px; border: 1px solid rgba(255,255,255,0.1);">✓ 16:9 Widescreen</span>
                <span style="font-size: 12px; color: #CBD5E1; background: rgba(255,255,255,0.06); padding: 8px 16px; border-radius: 8px; border: 1px solid rgba(255,255,255,0.1);">✓ 100% Vector Shapes</span>
                <span style="font-size: 12px; color: #CBD5E1; background: rgba(255,255,255,0.06); padding: 8px 16px; border-radius: 8px; border: 1px solid rgba(255,255,255,0.1);">✓ Full PPTX Editability</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Section 4: Final CTA & Footer
    st.markdown(
        """
        <div style="text-align: center; padding: 56px 20px 24px 20px;">
            <div class="section-eyebrow">Start Building</div>
            <h2 style="font-size: 36px; font-weight: 800; letter-spacing: -0.03em; color: #FFFFFF; margin: 10px 0 16px 0;">
                Your next presentation starts with one idea.
            </h2>
            <p style="font-size: 16px; color: #94A3B8; margin-bottom: 28px;">
                Scroll up to the composer and generate your presentation in seconds.
            </p>
        </div>

        <div class="footer-wrap">
            <div>AI PPT Maker Studio • Powered by Google Gemini & Python-PPTX</div>
            <div>
                <a href="https://github.com/rishalck/ai-ppt-maker" target="_blank" style="color: #94A3B8; text-decoration: none; margin-left: 16px;">GitHub Repository</a>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ============================================================================
# ENTRY POINT
# ============================================================================

def running_in_streamlit() -> bool:
    """Return True when the script is being executed by the Streamlit runtime."""
    return bool(
        "streamlit" in sys.modules
        or os.getenv("STREAMLIT_SERVER_PORT")
        or os.getenv("STREAMLIT_SERVER_ADDRESS")
        or os.getenv("STREAMLIT_GLOBAL")
        or os.getenv("STREAMLIT_APP")
    )


def main():
    if running_in_streamlit():
        streamlit_app()
    else:
        cli()


if __name__ == "__main__":
    main()
