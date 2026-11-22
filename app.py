"""
app.py - Context-Aware Multi-Input PPTX Generator

A Streamlit web application that creates stunning PowerPoint presentations
from multiple input sources: YouTube videos, PDFs, Word documents, existing
PPTX files, and custom prompts.

Supports:
- Multi-Provider LLMs: Google Gemini, OpenRouter, OpenAI, Groq, Ollama
- 2 Generation Modes:
    1. Procedural Canvas: Draws custom layouts with doodle characters & bento cards.
    2. Template Infill: Populates pre-designed templates from templates/ or user upload.
"""

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
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Context-Aware PPTX Generator",
    page_icon="🎨",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------------------
# Custom CSS for crisp high contrast in both dark & light modes
# ---------------------------------------------------------------------------
st.markdown("""
<style>
    .section-card {
        background: rgba(255, 255, 255, 0.05);
        border: 1px solid rgba(128, 128, 128, 0.2);
        border-radius: 12px;
        padding: 20px;
        margin-bottom: 16px;
    }
    .status-pill {
        display: inline-block;
        padding: 4px 12px;
        border-radius: 20px;
        font-size: 13px;
        font-weight: 600;
    }
    .status-success { background: #166534; color: #DCFCE7; }
    .status-error   { background: #991B1B; color: #FEE2E2; }
    .status-info    { background: #1E40AF; color: #DBEAFE; }
    .status-warning { background: #854D0E; color: #FEF9C3; }
    
    .hero {
        text-align: center;
        padding: 10px 0 20px 0;
    }
    .hero h1 {
        font-size: 2.2rem;
        font-weight: 800;
        margin-bottom: 4px;
    }
    .hero p {
        opacity: 0.8;
        font-size: 1.05rem;
    }
</style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Sidebar: AI Provider, Model & Mode Configuration
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown("## ⚙️ AI Provider & Settings")
    st.markdown("---")

    # 1. Provider Selector
    provider_choice = st.selectbox(
        "🧠 AI Provider",
        options=["Google Gemini", "OpenRouter", "OpenAI", "Groq / Custom"],
        index=0,
        help="Select which AI service to use for content synthesis.",
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
            "🤖 Gemini Model",
            options=["gemini-3.5-flash", "gemini-3.5-flash-lite", "gemini-3.8-flash"],
            index=0,
            help="gemini-3.5-flash is fast, reliable, and verified on your key.",
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
            help="OpenAI-compatible endpoint URL (e.g., http://localhost:11434/v1 for Ollama).",
        )
        api_key = st.text_input(
            "🔑 API Key",
            value=env_groq if env_groq else "",
            type="password",
            help="Leave as 'ollama' or any string if using local Ollama.",
        )
        model_choice = st.text_input(
            "🤖 Model Name",
            value="llama-3.3-70b-versatile",
            help="Exact model identifier on your provider.",
        )

    st.markdown("---")

    # 3. Generation Mode Selector
    st.markdown("## 🎨 Presentation Style & Mode")
    generation_mode = st.radio(
        "Generation Engine Mode",
        options=["Procedural Canvas (Dynamic Shapes & Characters)", "Template Infill (Clone .pptx Template)"],
        index=0,
        help="Procedural draws custom slides from scratch with characters and bento cards. Template Infill clones an existing .pptx design.",
    )

    if generation_mode.startswith("Procedural"):
        theme_choice = st.selectbox(
            "Visual Theme",
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
        # Template Infill mode: list templates from templates/
        available_templates = list_available_templates()
        template_names = [t["name"] for t in available_templates] + ["Upload Custom .pptx Template..."]
        chosen_tpl_name = st.selectbox("Choose Template from templates/", options=template_names, index=0)

        if chosen_tpl_name == "Upload Custom .pptx Template...":
            tpl_upload = st.file_uploader("Upload custom .pptx template", type=["pptx"])
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
    st.markdown(
        "💡 **BYOK Security:** Your API key is stored only in your active browser session memory and is never saved permanently."
    )


# ---------------------------------------------------------------------------
# Main area
# ---------------------------------------------------------------------------
st.markdown(
    '<div class="hero">'
    '<h1>🎨 Context-Aware Multi-Input PPTX Generator</h1>'
    '<p>Synthesizes YouTube videos, documents, and prompts into high-impact presentations</p>'
    '</div>',
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Input Section
# ---------------------------------------------------------------------------
col_left, col_right = st.columns([1, 1], gap="large")

with col_left:
    st.markdown("### 📝 Prompt & Instructions")
    user_prompt = st.text_area(
        "What should this presentation be about?",
        height=120,
        placeholder=(
            "e.g., Create an executive summary focusing on the key "
            "architectural insights and takeaways from the input. "
            "Use the doodle character to explain each concept."
        ),
    )

    st.markdown("### 🔗 YouTube Links")
    youtube_urls = st.text_area(
        "Paste YouTube video URLs (one per line)",
        height=80,
        placeholder="https://youtu.be/fkxw0SrwVkA\nhttps://www.youtube.com/watch?v=...",
    )

with col_right:
    st.markdown("### 📄 Upload Documents")
    uploaded_files = st.file_uploader(
        "Upload PDF, DOCX, or PPTX reference files",
        type=["pdf", "docx", "pptx"],
        accept_multiple_files=True,
        help="Drag and drop or click to browse. Multiple reference files supported.",
    )

    if uploaded_files:
        st.markdown(
            f'<span class="status-pill status-success">'
            f'{len(uploaded_files)} file(s) attached</span>',
            unsafe_allow_html=True,
        )
        for f in uploaded_files:
            size_kb = len(f.getvalue()) / 1024
            st.markdown(f"  📎 **{f.name}** ({size_kb:.1f} KB)")

# ---------------------------------------------------------------------------
# Asset & Graphics Explorer (2,450+ vector graphics)
# ---------------------------------------------------------------------------
with st.expander("🖼️ Explore Graphic & Icon Asset Library (2,450+ Vectors Available)", expanded=False):
    catalog_path = PROJECT_ROOT / "assets" / "catalog.json"
    if catalog_path.exists():
        import json
        with open(catalog_path, "r") as cf:
            catalog_data = json.load(cf)

        cat_col1, cat_col2 = st.columns([1, 2])
        with cat_col1:
            categories = ["All Categories", "Character Illustrations"] + sorted(list(catalog_data.get("categories", {}).keys()))
            selected_cat = st.selectbox("Category", categories, index=0)
        with cat_col2:
            search_query = st.text_input("Search Assets", placeholder="e.g. rocket, brain, server, cloud, dance, money, security...")

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

        st.caption(f"Showing **{min(len(items), 24)}** of **{len(items)}** matching assets (Total library: **{catalog_data.get('total_assets', 0):,}** vector assets across 11 categories)")

        grid_cols = st.columns(4)
        for idx, item in enumerate(items[:24]):
            with grid_cols[idx % 4]:
                st.markdown(f"**`{item['name']}`**")
                st.caption(f"📁 {item.get('category')}")
                st.write(f"_{item.get('description', '')}_")

        st.info("💡 **Pro-Tip:** Mention any icon name (e.g., `icon='rocket'`, `icon='server'`) or character pose (e.g., `pose='meditating'`, `pose='dancing'`) in your prompt, and the AI will automatically include it!")

st.markdown("---")

# ---------------------------------------------------------------------------
# Generate Button
# ---------------------------------------------------------------------------
generate_col1, generate_col2, generate_col3 = st.columns([1, 2, 1])
with generate_col2:
    generate_btn = st.button(
        "🚀  Generate Presentation",
        type="primary",
        use_container_width=True,
        disabled=not api_key,
    )

if not api_key:
    st.info(f"👈 Enter your {provider_choice} API key in the sidebar to get started.")

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

    with st.status("🚀 Processing inputs and building presentation...", expanded=True) as status_box:

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
        st.markdown("---")
        st.markdown("### 🎉 Your Presentation is Ready!")

        col_info, col_download = st.columns([2, 1])

        with col_info:
            st.markdown(
                f'<span class="status-pill status-success">SUCCESS</span> '
                f'&nbsp; **{result["slide_count"]} slides generated**',
                unsafe_allow_html=True,
            )
            st.markdown(f"**Engine Mode:** {generation_mode.split('(')[0].strip()}")
            st.markdown(f"**AI Provider:** {provider_choice} (`{model_choice}`)")

        with col_download:
            with open(result["output_path"], "rb") as f:
                st.download_button(
                    label="⬇️  Download PPTX",
                    data=f.read(),
                    file_name="presentation.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    type="primary",
                    use_container_width=True,
                )

        if result.get("code"):
            with st.expander("🔍 View Generated Slide Python Code", expanded=False):
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
# Footer
# ---------------------------------------------------------------------------
st.markdown("---")
st.markdown(
    '<div style="text-align:center; opacity:0.6; font-size:13px; padding:10px;">'
    "Context-Aware PPTX Generator • Supports Google Gemini, OpenRouter, OpenAI & Custom Templates"
    "</div>",
    unsafe_allow_html=True,
)
