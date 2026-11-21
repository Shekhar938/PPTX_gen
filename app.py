"""
app.py - Context-Aware Multi-Input PPTX Generator

A Streamlit web application that creates stunning PowerPoint presentations
from multiple input sources: YouTube videos, PDFs, Word documents, existing
PPTX files, and custom prompts.

Usage:
    python3 -m streamlit run app.py
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
    /* Card containers with adaptive styling */
    .section-card {
        background: rgba(255, 255, 255, 0.05);
        border: 1px solid rgba(128, 128, 128, 0.2);
        border-radius: 12px;
        padding: 20px;
        margin-bottom: 16px;
    }
    /* Status badges */
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
    
    /* Header title */
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
# Sidebar: API Key, Theme & Model Configuration
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown("## ⚙️ Configuration")
    st.markdown("---")

    # API Key
    env_key = os.getenv("GEMINI_API_KEY", "")
    api_key = st.text_input(
        "🔑 Gemini API Key",
        value=env_key if env_key and env_key != "YOUR_GEMINI_API_KEY_HERE" else "",
        type="password",
        help="Get your key at https://aistudio.google.com/apikey",
    )

    st.markdown("---")

    # Model selector - gemini-3.5-flash is verified active on this API key
    model_choice = st.selectbox(
        "🤖 Gemini Model",
        options=[
            "gemini-3.5-flash (Recommended - Fast & Verified)",
            "gemini-3.5-flash-lite (Ultra Fast)",
            "gemini-3.8-flash (Latest Preview)",
        ],
        index=0,
        help="gemini-3.5-flash is verified active on your API key with fast response times.",
    )
    # Extract raw model identifier
    raw_model = model_choice.split(" ")[0]

    st.markdown("---")

    # Theme selector
    theme_choice = st.selectbox(
        "🎨 Presentation Theme",
        options=["Black & White Doodle", "Dark Modern Tech", "Executive Light"],
        index=0,
        help="Visual style and aesthetic for the generated presentation.",
    )
    theme_map = {
        "Black & White Doodle": "bw_doodle",
        "Dark Modern Tech": "dark_modern",
        "Executive Light": "executive_light",
    }
    selected_theme = theme_map[theme_choice]

    st.markdown("---")
    st.markdown(
        "💡 **Tip:** Create a `.env` file in the project folder with:\n"
        "```bash\nGEMINI_API_KEY=AIza...\n```"
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
            "e.g., Create an 8-slide executive summary focusing on the key "
            "architectural points and pitfalls discussed in the video. "
            "Use the doodle character to guide the story."
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
        "Upload PDF, DOCX, or PPTX files",
        type=["pdf", "docx", "pptx"],
        accept_multiple_files=True,
        help="Drag and drop or click to browse. Multiple documents supported.",
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
    st.info("👈 Enter your Gemini API key in the sidebar to get started.")

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

    # Create an expandable live status container showing step-by-step progress
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
            status_box.write("📄 **Step 2:** Parsing attached documents...")
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

        # ---- STEP 3: Synthesize with Gemini & Build Slides ----
        status_box.write(f"🧠 **Step 3:** Synthesizing content with **{raw_model}**...")
        from generator.gemini_director import GeminiDirector

        director = GeminiDirector(api_key=api_key, model=raw_model)

        output_dir = tempfile.mkdtemp()
        output_path = os.path.join(output_dir, "presentation.pptx")

        # Live callback passes each update directly to the status box
        def on_status_update(message: str):
            status_box.write(f"  {message}")

        result = director.generate_and_build(
            user_prompt=user_prompt or "Create a comprehensive presentation from the provided content.",
            youtube_data=youtube_data if youtube_data else None,
            document_data=document_data if document_data else None,
            theme=selected_theme,
            output_path=output_path,
            max_retries=1,
            status_callback=on_status_update,
        )

        # Update status box verdict
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
            st.markdown(f"**Theme:** {theme_choice}")
            st.markdown(f"**Model Used:** {raw_model}")

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

        with st.expander("🔍 View Generated Slide Python Code", expanded=False):
            st.code(result.get("code", ""), language="python")

    else:
        st.error(f"Generation error: {result.get('error', 'Unknown error')}")
        st.info(
            "💡 **Troubleshooting Tips:**\n"
            "- If you see `503 UNAVAILABLE`: Google's free servers for that model are busy. Switch to **gemini-2.0-flash** in the sidebar.\n"
            "- If you see `403 / PERMISSION_DENIED`: Verify your API key at [Google AI Studio](https://aistudio.google.com/apikey)."
        )
        if result.get("code"):
            with st.expander("🔍 Generated Code (before error)", expanded=False):
                st.code(result.get("code", ""), language="python")


# ---------------------------------------------------------------------------
# Footer
# ---------------------------------------------------------------------------
st.markdown("---")
st.markdown(
    '<div style="text-align:center; opacity:0.6; font-size:13px; padding:10px;">'
    "Context-Aware PPTX Generator • Built with Google Gemini & python-pptx"
    "</div>",
    unsafe_allow_html=True,
)
