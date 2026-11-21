"""
unified_director.py - Multi-Provider LLM Orchestrator for Presentation Generation

Supports:
- Google Gemini (native google-genai SDK)
- OpenRouter (Claude, Llama, DeepSeek via OpenAI-compatible endpoint)
- OpenAI (GPT-4o, GPT-4o-mini)
- Groq / Custom / Local Ollama (Any OpenAI-compatible endpoint)

Supports both:
1. Procedural Mode (Generates Python SlideBuilder code)
2. Template Infill Mode (Generates structured JSON mapped to template placeholders)
"""

from __future__ import annotations

import json
import logging
import os
import re
import textwrap
import time
import traceback
from typing import Any, Dict, List, Optional

logger = logging.getLogger(__name__)

# Standard system prompt for Procedural Python code generation
from .gemini_director import SYSTEM_PROMPT, _build_content_prompt

# System prompt for Template Infill JSON generation
TEMPLATE_SYSTEM_PROMPT = textwrap.dedent("""\
You are an expert presentation content designer. Your job is to take raw content
(YouTube transcripts, document text, user instructions) and produce a structured
slide deck matching a pre-designed PowerPoint template.

## OUTPUT FORMAT
You MUST output ONLY a valid JSON object matching this schema:
```json
{
  "slides": [
    {
      "slide_number": 1,
      "title": "Main Presentation Title",
      "subtitle": "Subtitle or Executive Summary description"
    },
    {
      "slide_number": 2,
      "title": "Section or Overview Title",
      "cards": [
        {"title": "Pillar 1", "body": "Concise key takeaway or metric"},
        {"title": "Pillar 2", "body": "Concise key takeaway or metric"},
        {"title": "Pillar 3", "body": "Concise key takeaway or metric"}
      ]
    }
  ]
}
```
Output ONLY the raw JSON. No markdown fences, no explanatory text.
""")

PROVIDER_DEFAULTS = {
    "Google Gemini": {
        "base_url": None,
        "default_model": "gemini-3.5-flash",
        "models": ["gemini-3.5-flash", "gemini-3.5-flash-lite", "gemini-3.8-flash"],
    },
    "OpenRouter": {
        "base_url": "https://openrouter.ai/api/v1",
        "default_model": "deepseek/deepseek-chat",
        "models": [
            "deepseek/deepseek-chat",
            "anthropic/claude-3.5-sonnet",
            "meta-llama/llama-3.3-70b-instruct",
            "google/gemini-2.0-flash-001",
        ],
    },
    "OpenAI": {
        "base_url": "https://api.openai.com/v1",
        "default_model": "gpt-4o-mini",
        "models": ["gpt-4o-mini", "gpt-4o"],
    },
    "Groq / Custom": {
        "base_url": "https://api.groq.com/openai/v1",
        "default_model": "llama-3.3-70b-versatile",
        "models": ["llama-3.3-70b-versatile", "mixtral-8x7b-32768"],
    },
}


class RobustNamespace(dict):
    """Execution namespace that automatically evaluates unknown identifier names as strings.
    For example, if the LLM writes icon=lightbulb or pose=thinking without quotes,
    it automatically resolves to 'lightbulb' or 'thinking' instead of raising NameError.
    """
    def __missing__(self, key):
        return key


# Common visual tokens pre-registered in the namespace
KNOWN_TOKENS = [
    "brain", "database", "shield", "lightning", "chart_bar", "target", "star",
    "arrow_up", "clock", "check_circle", "warning", "lightbulb", "users", "code", "globe",
    "thinking", "pointing", "celebrating", "confused", "presenting", "reading",
    "left", "right", "center", "full", "bar", "column", "pie", "line", "area",
]


class UnifiedDirector:
    """Multi-provider AI orchestrator supporting Gemini, OpenRouter, OpenAI, and custom endpoints."""

    def __init__(
        self,
        provider: str = "Google Gemini",
        api_key: str = "",
        model: str = "",
        base_url: Optional[str] = None,
    ):
        self.provider = provider
        self.api_key = api_key
        config = PROVIDER_DEFAULTS.get(provider, PROVIDER_DEFAULTS["Google Gemini"])
        self.model = model or config["default_model"]
        self.base_url = base_url or config["base_url"]

    def _call_llm(self, prompt: str, system_instruction: str, status_callback=None) -> str:
        """Call the selected AI provider with error handling and status updates."""
        if self.provider == "Google Gemini":
            from google import genai
            from google.genai import errors as genai_errors

            client = genai.Client(api_key=self.api_key)
            models_to_try = [self.model, "gemini-3.5-flash", "gemini-3.5-flash-lite", "gemini-3.8-flash"]
            # Deduplicate preserving order
            models_to_try = list(dict.fromkeys(models_to_try))

            for m in models_to_try:
                for attempt in range(3):
                    wait = 2 ** attempt
                    try:
                        if status_callback:
                            status_callback(f"Connecting to Gemini ({m}, attempt {attempt+1})...")
                        resp = client.models.generate_content(
                            model=m,
                            contents=prompt,
                            config=genai.types.GenerateContentConfig(
                                system_instruction=system_instruction,
                                temperature=0.7,
                                max_output_tokens=16000,
                            ),
                        )
                        if status_callback:
                            status_callback(f"✅ Generated output with {m}!")
                        return resp.text.strip()
                    except (genai_errors.ServerError, Exception) as exc:
                        if "503" in str(exc) or "UNAVAILABLE" in str(exc):
                            if status_callback:
                                status_callback(f"⏳ {m} high demand, retrying in {wait}s...")
                            time.sleep(wait)
                            continue
                        break
            raise RuntimeError(f"All Gemini models failed on provider {self.provider}.")

        else:
            # Standard OpenAI-compatible API (OpenRouter, OpenAI, Groq, Ollama)
            from openai import OpenAI

            client = OpenAI(api_key=self.api_key, base_url=self.base_url)
            if status_callback:
                status_callback(f"Connecting to {self.provider} ({self.model})...")

            response = client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": system_instruction},
                    {"role": "user", "content": prompt},
                ],
                temperature=0.7,
            )
            if status_callback:
                status_callback(f"✅ Received response from {self.provider} ({self.model})!")
            return response.choices[0].message.content.strip()

    def generate_procedural_slides(
        self,
        user_prompt: str,
        youtube_data: list[dict] | None = None,
        document_data: list[dict] | None = None,
        theme: str = "bw_doodle",
        output_path: str = "output.pptx",
        status_callback=None,
    ) -> dict:
        """Generate Python slide code and build the presentation using SlideCanvas.
        Includes automatic code repair and robust namespace handling for NameErrors.
        """
        content_prompt = _build_content_prompt(user_prompt, youtube_data, document_data)

        from .slide_builder import PresentationDeck

        last_error = None
        code = ""

        # Up to 2 attempts (initial generation + 1 auto-repair if syntax/runtime error occurs)
        for attempt in range(2):
            if attempt == 0:
                raw_code = self._call_llm(content_prompt, SYSTEM_PROMPT, status_callback=status_callback)
            else:
                if status_callback:
                    status_callback("🛠️ Auto-repairing slide code...")
                repair_prompt = (
                    f"{content_prompt}\n\n"
                    f"## CRITICAL ERROR IN PREVIOUS ATTEMPT\n"
                    f"The code produced this error: {last_error}\n\n"
                    f"Fix the code and output ONLY valid Python code using the SlideBuilder API."
                )
                raw_code = self._call_llm(repair_prompt, SYSTEM_PROMPT, status_callback=status_callback)

            # Strip markdown fences
            code = re.sub(r"^```(?:python)?\s*\n", "", raw_code)
            code = re.sub(r"\n```\s*$", "", code)

            if status_callback:
                status_callback("🎨 Rendering PowerPoint shapes, bento cards & vector SVGs...")

            deck = PresentationDeck(theme=theme)
            namespace = RobustNamespace({
                "deck": deck,
                "PresentationDeck": PresentationDeck,
                "__builtins__": __builtins__,
            })
            for token in KNOWN_TOKENS:
                namespace[token] = token

            try:
                exec(code, namespace)
                saved_path = deck.save(output_path)
                return {
                    "success": True,
                    "output_path": saved_path,
                    "slide_count": deck.slide_count,
                    "error": None,
                    "code": code,
                }
            except Exception as e:
                last_error = f"{type(e).__name__}: {e}"
                logger.warning("Slide execution attempt %d failed: %s", attempt + 1, last_error)
                if status_callback:
                    status_callback(f"⚠️ Code issue: {last_error}. Retrying...")

        return {
            "success": False,
            "output_path": None,
            "slide_count": 0,
            "error": f"Failed after auto-repair attempt: {last_error}",
            "code": code,
        }

    def generate_template_slides(
        self,
        template_path: str,
        user_prompt: str,
        youtube_data: list[dict] | None = None,
        document_data: list[dict] | None = None,
        output_path: str = "output.pptx",
        status_callback=None,
    ) -> dict:
        """Generate structured JSON and inject it into an existing PowerPoint template."""
        content_prompt = _build_content_prompt(user_prompt, youtube_data, document_data)
        content_prompt += "\n\nMap this content into slides with title, subtitle, and cards."

        raw_json = self._call_llm(content_prompt, TEMPLATE_SYSTEM_PROMPT, status_callback=status_callback)

        # Clean JSON fences if present
        clean_json_str = re.sub(r"^```(?:json)?\s*\n", "", raw_json)
        clean_json_str = re.sub(r"\n```\s*$", "", clean_json_str).strip()

        try:
            data = json.loads(clean_json_str)
            slides_data = data.get("slides", [])
        except Exception:
            # Fallback regex extraction if raw json was surrounded by text
            m = re.search(r"\{.*\}", clean_json_str, re.DOTALL)
            if m:
                data = json.loads(m.group(0))
                slides_data = data.get("slides", [])
            else:
                raise ValueError("Could not parse structured slide JSON from LLM output.")

        if status_callback:
            status_callback(f"📄 Injecting content into template '{os.path.basename(template_path)}'...")

        from .template_engine import apply_content_to_template
        return apply_content_to_template(template_path, slides_data, output_path)
