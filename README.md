# 🎨 Context-Aware Multi-Input PPTX Generator

Transform YouTube videos, PDFs, Word documents, and ideas into stunning, professional PowerPoint presentations using Google Gemini AI and OpenAi Compitable LLM api.

## ✨ Features

- 🎬 **YouTube Transcript Extraction** — Paste any YouTube URL; captions are extracted automatically
- 📄 **Multi-Document Ingestion** — PDF, DOCX, and PPTX files all parsed and unified
- 🧠 **Gemini-Powered Synthesis** — All content fused into a coherent narrative
- 🎨 **3 Visual Themes** — Black & White Doodle, Dark Modern Tech, Executive Light
- 🖼️ **21 Built-in SVG Assets** — 6 character poses + 15 line icons, all vector-sharp
- 📊 **Native PowerPoint Charts** — Fully editable in PowerPoint/Google Slides
- 🔄 **Auto-Retry on Failure** — If generated code fails, Gemini fixes and retries

## 🚀 Quick Start

### 1. Set your API key

Copy `.env.example` to `.env` and add your Gemini API key:
```bash
cp .env.example .env
# Edit .env and replace YOUR_GEMINI_API_KEY_HERE with your actual key
# Get one free at: https://aistudio.google.com/apikey
```

### 2. Install dependencies (first time only)

```bash
python3 -m pip install --user --break-system-packages -r requirements.txt
```

### 3. Launch the app

```bash
python3 -m streamlit run app.py
```

The app opens at **http://localhost:8501**

---

## 📁 Project Structure

```
PPTX_gen/
├── app.py                      # Streamlit web UI
├── .env.example                # API key template
├── requirements.txt            # Python dependencies
│
├── assets/
│   ├── characters/             # 6 black-and-white character SVG poses
│   │   ├── thinking.svg
│   │   ├── pointing.svg
│   │   ├── celebrating.svg
│   │   ├── confused.svg
│   │   ├── presenting.svg
│   │   └── reading.svg
│   └── icons/                  # 15 minimalist line icons (SVG)
│       └── brain.svg, database.svg, ...
│
├── generator/
│   ├── theme.py                # Color palettes, typography, layout constants
│   ├── primitives.py           # Low-level python-pptx drawing functions
│   ├── slide_builder.py        # High-level SlideCanvas / PresentationDeck API
│   └── gemini_director.py      # LLM orchestrator — generates + executes slide code
│
├── parsers/
│   ├── youtube_parser.py       # YouTube transcript extractor
│   └── document_parser.py      # PDF, DOCX, PPTX parsers
│
└── utils/
    └── helpers.py              # Text sanitization, number extraction, paths
```

---

## 🎨 Available Themes

| Theme | Style | Best For |
|---|---|---|
| **Black & White Doodle** | Cream bg, charcoal, doodle characters | Casual, educational, viral-style |
| **Dark Modern Tech** | Deep slate, blue/violet accents | Tech talks, dev presentations |
| **Executive Light** | Clean white, royal blue | Corporate, investor decks |

---

## 🛠️ How It Works

```
YouTube URLs ──┐
PDF / DOCX    ──┼──► Unified Ingestion ──► Gemini Synthesis ──► SlideBuilder API ──► .pptx
PPTX files    ──┘       (parsers/)        (gemini_director)    (slide_builder)
User Prompt   ──┘
```

1. **Ingest** — All inputs are converted to clean structured text
2. **Synthesize** — Gemini reads everything and writes Python code using the SlideBuilder API
3. **Execute** — The code runs locally to build pixel-perfect slides
4. **Download** — Click to get a fully editable `.pptx` file

---

## 🎯 Slide Components Available to the AI

The AI can build any combination of these per slide:

- `add_hero_heading()` — Full-width title with badge and subtitle
- `add_character()` — Doodle character (thinking/pointing/celebrating/confused/presenting/reading)
- `add_bento_cards()` — Grid of styled content cards with icons
- `add_stat_banner()` — Row of large-number metric callouts
- `add_bullet_list()` — Styled bullet points
- `add_comparison()` — Side-by-side comparison layout
- `add_timeline()` — Horizontal process flow steps
- `add_quote()` — Large pull-quote callout
- `add_chart()` — Native editable bar/column/line/pie chart
- `add_table()` — Styled alternating-row table
- `add_content_with_image()` — Split text + image layout
- `add_footer()` — Slide footer with page numbers
