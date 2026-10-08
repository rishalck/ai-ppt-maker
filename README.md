# 🎯 AI PowerPoint Maker

An automated AI-powered PowerPoint presentation generator powered by Google Gemini and Python.

## 🚀 Features

- **Gemini-Powered Content Generation**: Intelligent presentation planning, structuring, and slide generation.
- **Multiple Visual Themes**: Dark, Light, Corporate, Neon, Minimal, and more.
- **Rich Slide Layouts**: Title slides, bullet points, KPI metrics, tables, timelines, comparisons, and quote slides.
- **Charts & Visuals**: Automatic chart generation using `matplotlib` and contextual image embedding.
- **Dual Interface**:
  - **Web UI**: Interactive dashboard built with Streamlit.
  - **CLI Mode**: Terminal-based generator for quick exports.
- **Speaker Notes & Auto Text-Fitting**: Formatted for presentation readiness.

---

## 🛠️ Installation

1. **Clone the repository**:
   ```bash
   git clone https://github.com/rishalck/ai-ppt-maker.git
   cd ai-ppt-maker
   ```

2. **Create and activate a virtual environment**:
   ```bash
   python -m venv .venv
   # Windows:
   .venv\Scripts\activate
   # macOS/Linux:
   source .venv/bin/activate
   ```

3. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

4. **Set up your API Key**:
   Copy `.env.example` to `.env` and add your Gemini API Key:
   ```bash
   cp .env.example .env
   ```
   Edit `.env`:
   ```env
   GEMINI_API_KEY="your_gemini_api_key_here"
   ```

---

## 💻 Usage

### Streamlit Web UI
```bash
streamlit run ppt_maker.py
```

### CLI Mode
```bash
python ppt_maker.py
```

---

## 📦 Requirements

- `google-genai`
- `python-pptx`
- `pillow`
- `requests`
- `matplotlib`
- `python-dotenv`
- `streamlit`

---

## 📄 License

MIT License
