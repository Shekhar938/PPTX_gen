"""
app.py - Context-Aware Multi-Input PPTX Generator

A modern, royal-themed Streamlit application that transforms YouTube videos,
documents, and ideas into executive-ready PowerPoint presentations.

Supports:
- Multi-Provider LLMs: Google Gemini, OpenRouter, OpenAI, Groq, Ollama
- 2 Generation Modes:
    1. Procedural Canvas: Dynamic shapes, bento cards, charts & vector doodles.
    2. Template Infill: Populates pre-designed templates from templates/ or user upload.
"""

import base64
import json
import os
import sys
import tempfile
from pathlib import Path

import streamlit as st
from dotenv import load_dotenv

# ---------------------------------------------------------------------------
# Project path setup
# ---------------------------------------------------------------------------
PROJECT_ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(PROJECT_ROOT))

load_dotenv(PROJECT_ROOT / ".env")

from generator.unified_director import UnifiedDirector, PROVIDER_DEFAULTS
from generator.template_engine import list_available_templates

# ---------------------------------------------------------------------------
# Page configuration
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="PPTX Generator — Royal Executive Suite",
    page_icon="👑",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------------------
# Helper: Embed project vector SVG assets as base64 Data URIs
# ---------------------------------------------------------------------------
@st.cache_data
def get_svg_uri(rel_path: str) -> str:
    """Read a project SVG asset and return a browser-renderable data URI."""
    try:
        p = PROJECT_ROOT / rel_path
        if p.exists():
            b64 = base64.b64encode(p.read_bytes()).decode("utf-8")
            return f"data:image/svg+xml;base64,{b64}"
    except Exception:
        pass
    return ""


# Mascot and UI vector assets
MASCOT_PRESENTING = get_svg_uri("assets/characters/presenting.svg")
MASCOT_CELEBRATING = get_svg_uri("assets/characters/celebrating.svg")
MASCOT_MEDITATING = get_svg_uri("assets/characters/meditating.svg")
ICON_SPARKLES = get_svg_uri("assets/icons/sparkles.svg")
ICON_VIDEO = get_svg_uri("assets/icons/video.svg")
ICON_FILE = get_svg_uri("assets/icons/file-text.svg")
ICON_ROCKET = get_svg_uri("assets/icons/rocket.svg")

# ---------------------------------------------------------------------------
# Royal Glassmorphism Design System (Modern CSS)
# ---------------------------------------------------------------------------
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;600&display=swap');

    /* Global typography */
    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
    }

    /* Ambient luxury background */
    .stApp {
        background: radial-gradient(circle at 15% 10%, rgba(99, 102, 241, 0.12) 0%, transparent 40%),
                    radial-gradient(circle at 85% 20%, rgba(168, 85, 247, 0.10) 0%, transparent 45%),
                    radial-gradient(circle at 50% 90%, rgba(245, 158, 11, 0.05) 0%, transparent 50%),
                    #080A12;
        color: #F8FAFC;
    }

    /* Sidebar luxury glass */
    [data-testid="stSidebar"] {
        background: rgba(12, 16, 29, 0.85) !important;
        backdrop-filter: blur(20px);
        border-right: 1px solid rgba(255, 255, 255, 0.06);
    }
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h2,
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h3 {
        color: #F8FAFC !important;
        font-weight: 700;
        letter-spacing: -0.01em;
    }

    /* Hero Banner */
    .royal-hero {
        text-align: center;
        padding: 30px 20px 20px 20px;
        position: relative;
    }
    .royal-badge {
        display: inline-flex;
        align-items: center;
        gap: 8px;
        padding: 6px 16px;
        background: linear-gradient(135deg, rgba(99, 102, 241, 0.2) 0%, rgba(168, 85, 247, 0.2) 100%);
        border: 1px solid rgba(168, 85, 247, 0.4);
        border-radius: 9999px;
        font-size: 12px;
        font-weight: 700;
        letter-spacing: 0.08em;
        text-transform: uppercase;
        color: #C084FC;
        margin-bottom: 16px;
        box-shadow: 0 0 25px rgba(168, 85, 247, 0.2);
    }
    .royal-title {
        font-size: 2.8rem;
        font-weight: 800;
        letter-spacing: -0.03em;
        background: linear-gradient(135deg, #FFFFFF 20%, #E2E8F0 50%, #A855F7 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 8px;
        line-height: 1.15;
    }
    .royal-subtitle {
        color: #94A3B8;
        font-size: 1.15rem;
        max-width: 720px;
        margin: 0 auto 20px auto;
        line-height: 1.5;
        font-weight: 400;
    }

    /* Feature Pills */
    .pill-cluster {
        display: flex;
        justify-content: center;
        flex-wrap: wrap;
        gap: 10px;
        margin-bottom: 24px;
    }
    .feature-pill {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 6px 14px;
        background: rgba(255, 255, 255, 0.04);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 20px;
        font-size: 12.5px;
        color: #CBD5E1;
        font-weight: 500;
        backdrop-filter: blur(10px);
    }

    /* Glassmorphic input cards */
    .glass-card {
        background: rgba(17, 24, 39, 0.65);
        backdrop-filter: blur(16px);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 18px;
        padding: 24px;
        box-shadow: 0 10px 30px -10px rgba(0, 0, 0, 0.5);
        transition: transform 0.25s cubic-bezier(0.16, 1, 0.3, 1), border-color 0.25s ease, box-shadow 0.25s ease;
        margin-bottom: 20px;
    }
    .glass-card:hover {
        transform: translateY(-2px);
        border-color: rgba(139, 92, 246, 0.35);
        box-shadow: 0 15px 35px -5px rgba(99, 102, 241, 0.2);
    }
    .card-header-row {
        display: flex;
        align-items: center;
        gap: 10px;
        margin-bottom: 12px;
    }
    .card-header-row h3 {
        margin: 0;
        font-size: 1.15rem;
        font-weight: 700;
        color: #F8FAFC;
    }

    /* Status indicators */
    .status-pill {
        display: inline-block;
        padding: 5px 14px;
        border-radius: 20px;
        font-size: 12.5px;
        font-weight: 700;
        letter-spacing: 0.03em;
    }
    .status-success { background: rgba(34, 197, 94, 0.15); color: #4ADE80; border: 1px solid rgba(74, 222, 128, 0.3); }
    .status-error   { background: rgba(239, 68, 68, 0.15); color: #F87171; border: 1px solid rgba(248, 113, 113, 0.3); }
    .status-info    { background: rgba(99, 102, 241, 0.15); color: #818CF8; border: 1px solid rgba(129, 140, 248, 0.3); }

    /* Custom Input Fields */
    .stTextArea textarea, .stTextInput input {
        background: rgba(15, 23, 42, 0.75) !important;
        border: 1px solid rgba(255, 255, 255, 0.1) !important;
        border-radius: 12px !important;
        color: #F8FAFC !important;
        font-size: 14px !important;
        transition: all 0.2s ease !important;
    }
    .stTextArea textarea:focus, .stTextInput input:focus {
        border-color: #8B5CF6 !important;
        box-shadow: 0 0 15px rgba(139, 92, 246, 0.3) !important;
    }

    /* Primary Royal Button */
    div.stButton > button:first-child {
        background: linear-gradient(135deg, #6366F1 0%, #8B5CF6 50%, #A855F7 100%) !important;
        color: #FFFFFF !important;
        border: 1px solid rgba(255, 255, 255, 0.2) !important;
        border-radius: 14px !important;
        padding: 14px 28px !important;
        font-size: 16px !important;
        font-weight: 700 !important;
        letter-spacing: 0.02em !important;
        box-shadow: 0 6px 25px rgba(139, 92, 246, 0.45) !important;
        transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1) !important;
    }
    div.stButton > button:first-child:hover {
        transform: translateY(-2px) scale(1.01) !important;
        box-shadow: 0 10px 30px rgba(168, 85, 247, 0.6) !important;
    }

    /* Asset Grid Tile */
    .asset-tile {
        background: rgba(15, 23, 42, 0.6);
        border: 1px solid rgba(255, 255, 255, 0.06);
        border-radius: 12px;
        padding: 12px;
        text-align: center;
        transition: all 0.2s ease;
        height: 100%;
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
    }
    .asset-tile:hover {
        border-color: rgba(168, 85, 247, 0.4);
        background: rgba(30, 41, 59, 0.6);
        transform: translateY(-2px);
    }
    .asset-svg-preview {
        width: 32px;
        height: 32px;
        filter: invert(1) drop-shadow(0 0 4px rgba(168, 85, 247, 0.4));
        margin-bottom: 6px;
    }
</style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Sidebar: AI Provider, Model & Mode Configuration
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown("""
    <div style="display:flex; align-items:center; gap:12px; margin-bottom:16px;">
        <span style="font-size:26px;">👑</span>
        <div>
            <div style="font-weight:800; font-size:16px; color:#F8FAFC;">Royal Deck Engine</div>
            <div style="font-size:11px; color:#A855F7; font-weight:600; text-transform:uppercase; letter-spacing:0.06em;">Executive AI Suite</div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("---")

    # 1. Provider Selector
    st.markdown("### 🧠 Intelligence Engine")
    provider_choice = st.selectbox(
        "AI Provider",
        options=["Google Gemini", "OpenRouter", "OpenAI", "Groq / Custom"],
        index=0,
        help="Select which AI reasoning provider to orchestrate your presentation.",
    )

    # 2. Dynamic API Key and Model Inputs based on Provider
    base_url = None
    if provider_choice == "Google Gemini":
        env_gemini = os.getenv("GEMINI_API_KEY", "")
        api_key = st.text_input(
            "🔑 Gemini API Key",
            value=env_gemini if env_gemini and not env_gemini.startswith("YOUR_") else "",
            type="password",
            help="Get your key at https://aistudio.google.com/apikey",
        )
        model_choice = st.selectbox(
            "🤖 Model",
            options=["gemini-3.5-flash", "gemini-3.5-flash-lite", "gemini-3.8-flash"],
            index=0,
            help="gemini-3.5-flash is verified active with rapid inference.",
        )

    elif provider_choice == "OpenRouter":
        env_or = os.getenv("OPENROUTER_API_KEY", "")
        api_key = st.text_input(
            "🔑 OpenRouter API Key",
            value=env_or if env_or else "",
            type="password",
            help="Get your key at https://openrouter.ai/keys",
        )
        model_choice = st.selectbox(
            "🤖 Model",
            options=[
                "deepseek/deepseek-chat",
                "anthropic/claude-3.5-sonnet",
                "meta-llama/llama-3.3-70b-instruct",
                "google/gemini-2.0-flash-001",
            ],
            index=0,
        )
        base_url = "https://openrouter.ai/api/v1"

    elif provider_choice == "OpenAI":
        env_oai = os.getenv("OPENAI_API_KEY", "")
        api_key = st.text_input(
            "🔑 OpenAI API Key",
            value=env_oai if env_oai else "",
            type="password",
            help="Get your key at https://platform.openai.com/api-keys",
        )
        model_choice = st.selectbox(
            "🤖 Model",
            options=["gpt-4o-mini", "gpt-4o"],
            index=0,
        )
        base_url = "https://api.openai.com/v1"

    else:  # Groq / Custom / Local Ollama
        env_groq = os.getenv("GROQ_API_KEY", "")
        base_url = st.text_input(
            "🌐 Base URL",
            value="https://api.groq.com/openai/v1",
            help="OpenAI-compatible endpoint URL (e.g. http://localhost:11434/v1).",
        )
        api_key = st.text_input(
            "🔑 API Key",
            value=env_groq if env_groq else "",
            type="password",
            help="API Key for this provider.",
        )
        model_choice = st.text_input(
            "🤖 Model Name",
            value="llama-3.3-70b-versatile",
            help="Exact model name.",
        )

    st.markdown("---")

    # 3. Generation Mode Selector
    st.markdown("### 🎨 Presentation Architecture")
    generation_mode = st.radio(
        "Engine Mode",
        options=["Procedural Canvas (Dynamic Visuals)", "Template Infill (Clone .pptx)"],
        index=0,
        help="Procedural draws custom slides with doodle characters and bento cards. Template Infill clones an existing .pptx design.",
    )

    if generation_mode.startswith("Procedural"):
        theme_choice = st.selectbox(
            "Aesthetic Theme",
            options=["Black & White Doodle", "Dark Modern Tech", "Executive Light"],
            index=0,
        )
        theme_map = {
            "Black & White Doodle": "bw_doodle",
            "Dark Modern Tech": "dark_modern",
            "Executive Light": "executive_light",
        }
        selected_theme = theme_map[theme_choice]
        selected_template_path = None
    else:
        available_templates = list_available_templates()
        template_names = [t["name"] for t in available_templates] + ["Upload Custom .pptx Template..."]
        chosen_tpl_name = st.selectbox("Select Template", options=template_names, index=0)

        if chosen_tpl_name == "Upload Custom .pptx Template...":
            tpl_upload = st.file_uploader("Upload .pptx template", type=["pptx"])
            if tpl_upload:
                tmp_tpl = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
                tmp_tpl.write(tpl_upload.getvalue())
                selected_template_path = tmp_tpl.name
            else:
                selected_template_path = None
        else:
            matching = [t for t in available_templates if t["name"] == chosen_tpl_name]
            selected_template_path = matching[0]["path"] if matching else None

    st.markdown("---")
    st.markdown("""
    <div style="font-size:12px; color:#64748B; line-height:1.4;">
        🔒 <b>Zero-Persistence BYOK:</b> Your API keys are retained exclusively in active browser session memory and never stored on disk.
    </div>
    """, unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Royal Hero Section
# ---------------------------------------------------------------------------
mascot_html = f'<img src="{MASCOT_PRESENTING}" style="width:72px; height:72px; filter:invert(1) drop-shadow(0 0 12px rgba(168,85,247,0.5));">' if MASCOT_PRESENTING else "🎨"

st.markdown(f"""
<div class="royal-hero">
    <div class="royal-badge">👑 Royal Presentation Suite • Next-Gen Deck Engine</div>
    <div style="margin-bottom:12px;">{mascot_html}</div>
    <div class="royal-title">Context-Aware PPTX Generator</div>
    <div class="royal-subtitle">
        Synthesize multi-modal YouTube lectures, documents, and research into polished, executive-ready presentation decks.
    </div>
    <div class="pill-cluster">
        <div class="feature-pill">⚡ 2,450+ Vector Graphics</div>
        <div class="feature-pill">🎬 YouTube Transcript Synthesis</div>
        <div class="feature-pill">📄 PDF, DOCX & PPTX Ingestion</div>
        <div class="feature-pill">🎨 16:9 Native PowerPoint Output</div>
        <div class="feature-pill">🍱 Auto-Balanced Bento Grids</div>
    </div>
</div>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Input Cards (Glassmorphism Layout)
# ---------------------------------------------------------------------------
col_left, col_right = st.columns([1, 1], gap="medium")

with col_left:
    st.markdown(f"""
    <div class="card-header-row">
        <img src="{ICON_SPARKLES}" style="width:20px; height:20px; filter:invert(1); opacity:0.8;">
        <h3>Creative Brief & Prompt</h3>
    </div>
    """, unsafe_allow_html=True)
    user_prompt = st.text_area(
        "Instructions & Focus",
        label_visibility="collapsed",
        height=120,
        placeholder="e.g. Create an 8-slide executive briefing analyzing key architectural tradeoffs and takeaways. Use the doodle character to guide the narrative.",
    )

    st.markdown(f"""
    <div class="card-header-row" style="margin-top:16px;">
        <img src="{ICON_VIDEO}" style="width:20px; height:20px; filter:invert(1); opacity:0.8;">
        <h3>YouTube Ingestion (URLs)</h3>
    </div>
    """, unsafe_allow_html=True)
    youtube_urls = st.text_area(
        "YouTube Video Links",
        label_visibility="collapsed",
        height=85,
        placeholder="https://youtu.be/fkxw0SrwVkA\nhttps://www.youtube.com/watch?v=...",
    )

with col_right:
    st.markdown(f"""
    <div class="card-header-row">
        <img src="{ICON_FILE}" style="width:20px; height:20px; filter:invert(1); opacity:0.8;">
        <h3>Reference Documents & Assets</h3>
    </div>
    """, unsafe_allow_html=True)
    uploaded_files = st.file_uploader(
        "Upload PDF, DOCX, or PPTX reference files",
        type=["pdf", "docx", "pptx"],
        accept_multiple_files=True,
        help="Drag and drop multiple reference documents. Text and structure will be extracted.",
    )

    if uploaded_files:
        st.markdown(
            f'<div style="margin-top:10px;"><span class="status-pill status-success">'
            f'✓ {len(uploaded_files)} file(s) attached</span></div>',
            unsafe_allow_html=True,
        )
        for f in uploaded_files:
            size_kb = len(f.getvalue()) / 1024
            st.markdown(f"<div style='font-size:12.5px; color:#94A3B8; margin-top:4px;'>📎 <b>{f.name}</b> ({size_kb:.1f} KB)</div>", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Royal Asset & Graphics Showroom (Visual Catalog with Real SVG Rendering)
# ---------------------------------------------------------------------------
st.markdown("<div style='height:16px;'></div>", unsafe_allow_html=True)

with st.expander("🖼️ Explore Visual Asset Library (2,450+ Vector SVGs with Descriptions)", expanded=False):
    catalog_path = PROJECT_ROOT / "assets" / "catalog.json"
    if catalog_path.exists():
        with open(catalog_path, "r") as cf:
            catalog_data = json.load(cf)

        filter_col1, filter_col2 = st.columns([1, 2])
        with filter_col1:
            categories = ["All Categories", "Character Illustrations"] + sorted(list(catalog_data.get("categories", {}).keys()))
            selected_cat = st.selectbox("Filter Category", categories, index=0)
        with filter_col2:
            search_query = st.text_input("Search Vector Assets", placeholder="e.g. rocket, brain, server, cloud, dance, money, security...")

        if selected_cat == "Character Illustrations":
            items = catalog_data.get("characters", [])
        elif selected_cat == "All Categories":
            items = catalog_data.get("characters", []) + catalog_data.get("icons", [])
        else:
            items = [i for i in catalog_data.get("icons", []) if i.get("category") == selected_cat]

        if search_query:
            sq = search_query.lower().strip()
            items = [
                i for i in items
                if sq in i.get("name", "").lower()
                or sq in i.get("description", "").lower()
                or any(sq in t.lower() for t in i.get("tags", []))
            ]

        st.markdown(f"<div style='font-size:12.5px; color:#A855F7; margin-bottom:12px; font-weight:600;'>Showing {min(len(items), 16)} of {len(items)} matching assets (Library Total: {catalog_data.get('total_assets', 0):,} vector graphics)</div>", unsafe_allow_html=True)

        # Show visual grid with real rendered SVG graphics
        grid_cols = st.columns(4)
        for idx, item in enumerate(items[:16]):
            with grid_cols[idx % 4]:
                svg_uri = get_svg_uri(item.get("path", ""))
                img_tag = f'<img class="asset-svg-preview" src="{svg_uri}">' if svg_uri else "🎨"
                st.markdown(f"""
                <div class="asset-tile">
                    {img_tag}
                    <div style="font-weight:700; font-size:13px; color:#F8FAFC; margin-bottom:2px;"><code>{item['name']}</code></div>
                    <div style="font-size:10px; color:#A855F7; text-transform:uppercase; font-weight:600;">{item.get('category')}</div>
                    <div style="font-size:11px; color:#94A3B8; margin-top:4px; line-height:1.2;">{item.get('description', '')[:50]}</div>
                </div>
                """, unsafe_allow_html=True)

        st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)
        st.info("💡 **Pro-Tip:** Mention any icon name (e.g. `icon='rocket'`, `icon='server'`) or character pose (e.g. `pose='meditating'`, `pose='dancing'`) in your prompt, and the AI will automatically include it!")


# ---------------------------------------------------------------------------
# Royal CTA Action Button
# ---------------------------------------------------------------------------
st.markdown("<div style='height:20px;'></div>", unsafe_allow_html=True)
_, btn_col, _ = st.columns([1, 2, 1])
with btn_col:
    generate_btn = st.button(
        "🚀  Synthesize & Generate Presentation",
        type="primary",
        use_container_width=True,
        disabled=not api_key,
    )

if not api_key:
    st.markdown(f"""
    <div style="text-align:center; padding:12px; color:#CBD5E1; font-size:13.5px;">
        👈 <i>Enter your {provider_choice} API key in the sidebar to activate generation.</i>
    </div>
    """, unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Generation Pipeline with Real-Time Live Status
# ---------------------------------------------------------------------------
if generate_btn and api_key:
    has_youtube = bool(youtube_urls and youtube_urls.strip())
    has_docs = bool(uploaded_files)
    has_prompt = bool(user_prompt and user_prompt.strip())

    if not has_youtube and not has_docs and not has_prompt:
        st.error("Please provide at least one input: a YouTube link, an uploaded document, or a prompt.")
        st.stop()

    if generation_mode.startswith("Template") and not selected_template_path:
        st.error("Please select or upload a valid .pptx template.")
        st.stop()

    with st.status("🚀 Orchestrating presentation synthesis...", expanded=True) as status_box:

        # ---- STEP 1: Parse YouTube transcripts ----
        youtube_data = []
        if has_youtube:
            status_box.write("🎬 **Step 1:** Extracting YouTube video transcripts...")
            from parsers.youtube_parser import extract_youtube_transcript

            urls = [u.strip() for u in youtube_urls.strip().split("\n") if u.strip()]
            for url in urls:
                try:
                    yt_result = extract_youtube_transcript(url)
                    youtube_data.append(yt_result)
                    vid_title = yt_result.get("title", "Unknown")
                    char_count = len(yt_result.get("transcript", ""))
                    status_box.write(f"  ✅ Extracted: **{vid_title}** ({char_count:,} characters)")
                except Exception as e:
                    status_box.write(f"  ⚠️ Could not extract transcript from {url}: {e}")

        # ---- STEP 2: Parse uploaded documents ----
        document_data = []
        if has_docs:
            status_box.write("📄 **Step 2:** Parsing attached reference documents...")
            from parsers.document_parser import parse_document

            for uploaded_file in uploaded_files:
                suffix = Path(uploaded_file.name).suffix
                with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                    tmp.write(uploaded_file.getvalue())
                    tmp_path = tmp.name

                try:
                    doc_result = parse_document(tmp_path)
                    doc_result["filename"] = uploaded_file.name
                    document_data.append(doc_result)
                    pages_or_slides = doc_result.get("pages") or doc_result.get("slide_count") or len(doc_result.get("sections", []))
                    status_box.write(f"  ✅ Parsed: **{uploaded_file.name}** ({pages_or_slides} segments/pages)")
                except Exception as e:
                    status_box.write(f"  ⚠️ Could not parse {uploaded_file.name}: {e}")
                finally:
                    if os.path.exists(tmp_path):
                        os.unlink(tmp_path)

        # ---- STEP 3: Unified AI Synthesis & Presentation Build ----
        status_box.write(f"🧠 **Step 3:** Synthesizing with **{provider_choice}** ({model_choice})...")
        director = UnifiedDirector(
            provider=provider_choice,
            api_key=api_key,
            model=model_choice,
            base_url=base_url,
        )

        output_dir = tempfile.mkdtemp()
        output_path = os.path.join(output_dir, "presentation.pptx")

        def on_status_update(message: str):
            status_box.write(f"  {message}")

        if generation_mode.startswith("Procedural"):
            result = director.generate_procedural_slides(
                user_prompt=user_prompt or "Create a comprehensive presentation from the provided content.",
                youtube_data=youtube_data if youtube_data else None,
                document_data=document_data if document_data else None,
                theme=selected_theme,
                output_path=output_path,
                status_callback=on_status_update,
            )
        else:
            result = director.generate_template_slides(
                template_path=selected_template_path,
                user_prompt=user_prompt or "Create a comprehensive presentation from the provided content.",
                youtube_data=youtube_data if youtube_data else None,
                document_data=document_data if document_data else None,
                output_path=output_path,
                status_callback=on_status_update,
            )

        if result["success"]:
            status_box.update(
                label=f"🎉 Presentation Ready! ({result['slide_count']} slides generated)",
                state="complete",
                expanded=False,
            )
        else:
            status_box.update(
                label="❌ Presentation Generation Failed",
                state="error",
                expanded=True,
            )

    # ---- Display Results ----
    if result["success"]:
        st.markdown("<div style='height:20px;'></div>", unsafe_allow_html=True)
        celebrate_img = f'<img src="{MASCOT_CELEBRATING}" style="width:64px; height:64px; filter:invert(1) drop-shadow(0 0 15px rgba(34,197,94,0.6));">' if MASCOT_CELEBRATING else "🎉"

        st.markdown(f"""
        <div class="glass-card" style="border-color: rgba(74, 222, 128, 0.4); text-align:center; padding:30px;">
            <div style="margin-bottom:8px;">{celebrate_img}</div>
            <h2 style="color:#FFFFFF; font-size:1.8rem; font-weight:800; margin-bottom:4px;">Your Executive Presentation is Ready</h2>
            <div style="color:#94A3B8; font-size:14px; margin-bottom:18px;">
                Generated <b>{result['slide_count']} slides</b> with native vector shapes, typography, and theme styling.
            </div>
            <div class="pill-cluster" style="margin-bottom:0;">
                <span class="status-pill status-success">SUCCESS</span>
                <div class="feature-pill">📊 {result['slide_count']} Slides</div>
                <div class="feature-pill">🤖 {provider_choice} ({model_choice})</div>
                <div class="feature-pill">🎨 {generation_mode.split('(')[0].strip()}</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        _, dl_col, _ = st.columns([1, 2, 1])
        with dl_col:
            with open(result["output_path"], "rb") as f:
                st.download_button(
                    label="⬇️  Download PowerPoint Presentation (.pptx)",
                    data=f.read(),
                    file_name="presentation.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    type="primary",
                    use_container_width=True,
                )

        if result.get("code"):
            with st.expander("🔍 Inspect Generated SlideBuilder Python Code", expanded=False):
                st.code(result.get("code", ""), language="python")

    else:
        st.error(f"Generation error: {result.get('error', 'Unknown error')}")
        st.info(
            "💡 **Troubleshooting Tips:**\n"
            "- Verify that your API key is valid for the selected provider.\n"
            "- If using OpenRouter or OpenAI, ensure your account has active credits.\n"
            "- For Google Gemini, gemini-3.5-flash is the verified high-performance model."
        )


# ---------------------------------------------------------------------------
# Royal Footer
# ---------------------------------------------------------------------------
st.markdown("<div style='height:40px;'></div>", unsafe_allow_html=True)
st.markdown(
    '<div style="text-align:center; color:#64748B; font-size:12.5px; padding:20px; border-top:1px solid rgba(255,255,255,0.06);">'
    "👑 <b>Royal Presentation Suite</b> • Powered by Google Gemini, OpenRouter & python-pptx • 2,450+ Vector Graphics"
    "</div>",
    unsafe_allow_html=True,
)
