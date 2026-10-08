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
            data["stats"].append({
                "number": clean_text(item.get("number")),
                "label": clean_text(item.get("label")),
            })

    for item in raw.get("timeline", []):
        if isinstance(item, dict):
            data["timeline"].append({
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
        page_title=APP_NAME,
        page_icon="📊",
        layout="wide",
    )

    st.markdown(
        """
        <style>
        .main-title {
            font-size: 48px;
            font-weight: 800;
            letter-spacing: -2px;
        }
        .subtitle {
            color: #8b93a7;
            font-size: 18px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        '<div class="main-title">AI PowerPoint Maker</div>',
        unsafe_allow_html=True,
    )

    st.markdown(
        '<div class="subtitle">Gemini + Python-PPTX • Generate polished presentations with pictures</div>',
        unsafe_allow_html=True,
    )

    st.divider()

    left, right = st.columns([1.0, 1.6])

    with left:
        topic = st.text_area(
            "What should the presentation be about?",
            placeholder="Example: Artificial Intelligence in Healthcare",
            height=130,
        )

        slides = st.slider(
            "Number of slides",
            min_value=3,
            max_value=30,
            value=10,
        )

        audience = st.selectbox(
            "Audience",
            [
                "College students",
                "School students",
                "Business audience",
                "Startup founders",
                "Technical professionals",
                "General audience",
            ],
        )

        style = st.selectbox(
            "Presentation style",
            [
                "Premium modern",
                "Minimal",
                "Corporate",
                "Academic",
                "Startup pitch",
                "Creative",
                "Futuristic",
            ],
        )

        theme = st.selectbox(
            "Theme",
            list(THEMES.keys()),
            index=0,
        )

        language = st.selectbox(
            "Language",
            [
                "English",
                "Malayalam",
                "Hindi",
                "Tamil",
                "Kannada",
            ],
        )

        include_images = st.checkbox(
            "Include pictures",
            value=True,
        )

        include_charts = st.checkbox(
            "Include charts",
            value=True,
        )

        include_notes = st.checkbox(
            "Include speaker notes",
            value=True,
        )

        author = st.text_input("Author")
        company = st.text_input("Company / college")

    with right:
        st.markdown("### What this generator creates")

        features = [
            "AI-generated slide structure",
            "Professional slide layouts",
            "Pictures on visual slides",
            "Comparison layouts",
            "Timeline layouts",
            "KPI / statistics cards",
            "Charts",
            "Tables",
            "Speaker notes",
            "Real .pptx export",
        ]

        for feature in features:
            st.markdown("✓ " + feature)

        st.info(
            "For images, this demo uses public Unsplash image URLs. "
            "For commercial use, connect a licensed image provider."
        )

    st.divider()

    generate = st.button(
        "✨ Generate PowerPoint",
        type="primary",
        use_container_width=True,
    )

    if generate:
        if not topic.strip():
            st.error("Enter a presentation topic.")
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

        progress = st.progress(0)

        try:
            progress.progress(10)
            st.write("Planning presentation...")

            engine = GeminiEngine()
            generated_slides = build_plan(engine, config)

            progress.progress(35)
            st.write("Improving content...")

            generated_slides = improve_slide_copy(
                engine,
                generated_slides,
            )

            progress.progress(50)

            if include_images:
                st.write("Preparing pictures...")
                generated_slides = improve_image_queries(
                    engine,
                    generated_slides,
                )

            progress.progress(70)
            st.write("Building PowerPoint...")

            builder = PPTXBuilder(config)
            output = builder.build(generated_slides)

            progress.progress(100)

            st.success(
                f"Created {len(generated_slides)} slides."
            )

            with open(output, "rb") as file:
                st.download_button(
                    "⬇️ Download PowerPoint",
                    data=file.read(),
                    file_name=output.name,
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "presentationml.presentation"
                    ),
                    use_container_width=True,
                )

            st.markdown("### Generated slides")

            for index, slide in enumerate(
                generated_slides,
                start=1,
            ):
                with st.expander(
                    f"{index:02d} — {slide.title}"
                ):
                    st.write(f"**Type:** {slide.type}")

                    if slide.subtitle:
                        st.write(slide.subtitle)

                    for bullet in slide.bullets:
                        st.write("• " + bullet)

        except Exception as exc:
            st.error(str(exc))
            st.code(traceback.format_exc())


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
