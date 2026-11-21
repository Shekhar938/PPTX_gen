"""
youtube_parser.py - YouTube Transcript Extractor for PPTX Generation.

Extracts transcripts and metadata from YouTube videos across multiple URL formats.
Uses youtube-transcript-api to fetch manual or auto-generated captions, prioritizing
English and falling back to any available language.
"""

from __future__ import annotations

import json
import logging
import re
import urllib.error
import urllib.parse
import urllib.request
from typing import Any, Dict, List, Optional

from youtube_transcript_api import (
    InvalidVideoId,
    NoTranscriptFound,
    TranscriptsDisabled,
    VideoUnavailable,
    VideoUnplayable,
    YouTubeTranscriptApi,
    YouTubeTranscriptApiException,
)

from utils.helpers import sanitize_text

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Custom Exceptions
# ---------------------------------------------------------------------------
class YouTubeParserError(Exception):
    """Base exception for YouTube parsing and transcript extraction errors."""
    pass


class VideoUnavailableError(YouTubeParserError):
    """Raised when a video is unavailable, deleted, private, or age-restricted."""
    pass


class TranscriptsDisabledError(YouTubeParserError):
    """Raised when subtitles/captions are disabled for the video."""
    pass


class TranscriptNotFoundError(YouTubeParserError):
    """Raised when no subtitles are found in the requested or fallback languages."""
    pass


# ---------------------------------------------------------------------------
# Video ID Extractor
# ---------------------------------------------------------------------------
def extract_video_id(url: str) -> str:
    """Extract the 11-character YouTube video ID from various URL formats.

    Supports:
    - Standard watch URL: ``https://www.youtube.com/watch?v=dQw4w9WgXcQ``
    - Shortened URL: ``https://youtu.be/dQw4w9WgXcQ``
    - Embedded URL: ``https://www.youtube.com/embed/dQw4w9WgXcQ``
    - Direct /v/ or /e/: ``https://www.youtube.com/v/dQw4w9WgXcQ``
    - Shorts URL: ``https://www.youtube.com/shorts/dQw4w9WgXcQ``
    - Live URL: ``https://www.youtube.com/live/dQw4w9WgXcQ``
    - Mobile URL: ``https://m.youtube.com/watch?v=dQw4w9WgXcQ``
    - Raw 11-character ID: ``dQw4w9WgXcQ``

    Args:
        url: YouTube video URL or raw video ID.

    Returns:
        str: 11-character YouTube video ID.

    Raises:
        ValueError: If a valid 11-character YouTube ID cannot be identified.
    """
    if not url or not isinstance(url, str):
        raise ValueError("URL must be a non-empty string.")

    cleaned_url = url.strip()

    # If it is already an 11-character alphanumeric ID
    if re.fullmatch(r"[a-zA-Z0-9_-]{11}", cleaned_url):
        return cleaned_url

    parsed = urllib.parse.urlparse(cleaned_url)
    hostname = (parsed.hostname or "").lower()
    path = parsed.path

    # Check query parameters (watch?v=...)
    if hostname in ("www.youtube.com", "youtube.com", "m.youtube.com"):
        if path == "/watch":
            query_params = urllib.parse.parse_qs(parsed.query)
            v_param = query_params.get("v")
            if v_param and re.fullmatch(r"[a-zA-Z0-9_-]{11}", v_param[0]):
                return v_param[0]

        # Check path prefixes like /embed/, /v/, /shorts/, /live/
        for prefix in ("/embed/", "/v/", "/shorts/", "/live/", "/e/"):
            if path.startswith(prefix):
                candidate = path[len(prefix):].split("/")[0].split("?")[0].split("&")[0]
                if re.fullmatch(r"[a-zA-Z0-9_-]{11}", candidate):
                    return candidate

    elif hostname in ("youtu.be", "www.youtu.be"):
        candidate = path.lstrip("/").split("/")[0].split("?")[0].split("&")[0]
        if re.fullmatch(r"[a-zA-Z0-9_-]{11}", candidate):
            return candidate

    # Robust regex fallback for any YouTube URL with an 11-char ID
    regex_pattern = (
        r"(?:v=|\/embed\/|\/v\/|\/shorts\/|\/live\/|\/e\/|youtu\.be\/)"
        r"([a-zA-Z0-9_-]{11})"
    )
    match = re.search(regex_pattern, cleaned_url)
    if match:
        return match.group(1)

    raise ValueError(
        f"Could not extract a valid YouTube video ID from '{url}'. "
        "Expected a valid YouTube URL or 11-character video ID."
    )


# ---------------------------------------------------------------------------
# Video Title Fetcher
# ---------------------------------------------------------------------------
def fetch_video_title(video_id: str) -> str:
    """Fetch the video title using YouTube's oEmbed endpoint.

    Falls back to a placeholder string if network is unavailable or request fails.

    Args:
        video_id: 11-character YouTube video ID.

    Returns:
        str: Video title or fallback title.
    """
    oembed_url = f"https://www.youtube.com/oembed?url=https://www.youtube.com/watch?v={video_id}&format=json"
    headers = {"User-Agent": "Mozilla/5.0 (PPTX Generator Agent)"}

    try:
        req = urllib.request.Request(oembed_url, headers=headers)
        with urllib.request.urlopen(req, timeout=5) as response:
            payload = json.loads(response.read().decode("utf-8"))
            title = payload.get("title")
            if title and isinstance(title, str):
                return sanitize_text(title)
    except Exception as err:
        logger.debug(f"Could not fetch video title via oEmbed for {video_id}: {err}")

    return f"YouTube Video ({video_id})"


# ---------------------------------------------------------------------------
# Transcript Extractor
# ---------------------------------------------------------------------------
def extract_youtube_transcript(url: str) -> Dict[str, Any]:
    """Extract transcript, metadata, and segmented timestamps from a YouTube video.

    Tries to retrieve English transcripts first (manual, then auto-generated),
    and falls back to any available language if English is not found.

    Args:
        url: Full YouTube video URL or 11-character video ID.

    Returns:
        Dict[str, Any]: Extracted data with keys:
            - 'video_id' (str): The 11-character YouTube video identifier.
            - 'title' (str): The video title.
            - 'transcript' (str): Complete concatenated transcript text.
            - 'segments' (List[Dict[str, Any]]): List of snippet dictionaries
              each with 'text' (str), 'start' (float seconds), and 'duration' (float seconds).

    Raises:
        ValueError: If the URL or video ID is malformed.
        TranscriptsDisabledError: If the video owner disabled transcripts.
        TranscriptNotFoundError: If no transcript exists in any language.
        VideoUnavailableError: If the video is private, deleted, or unavailable.
        YouTubeParserError: If an unexpected error occurs during transcript extraction.
    """
    video_id = extract_video_id(url)
    title = fetch_video_title(video_id)

    ytt_api = YouTubeTranscriptApi()
    transcript_obj = None

    preferred_languages = ["en", "en-US", "en-GB", "en-CA", "en-AU"]

    try:
        # Step 1: Attempt standard fetch prioritizing preferred English codes
        try:
            transcript_obj = ytt_api.fetch(video_id, languages=preferred_languages)
        except (NoTranscriptFound, Exception):
            pass

        # Step 2: If English was not found by fetch(), query the transcript list
        if transcript_obj is None:
            transcript_list = ytt_api.list(video_id)

            # Check explicitly for auto-generated English
            try:
                gen_transcript = transcript_list.find_generated_transcript(preferred_languages)
                transcript_obj = gen_transcript.fetch()
            except (NoTranscriptFound, Exception):
                pass

            # Step 3: Fall back to ANY available transcript (manual first, then auto-generated)
            if transcript_obj is None:
                for candidate_transcript in transcript_list:
                    transcript_obj = candidate_transcript.fetch()
                    break

        if transcript_obj is None:
            raise NoTranscriptFound(video_id, preferred_languages, None)

        # Parse snippets from the fetched transcript object
        # Snippets support .text, .start, .duration
        segments: List[Dict[str, Any]] = []
        for snippet in transcript_obj:
            text_cleaned = sanitize_text(snippet.text)
            if text_cleaned:
                segments.append({
                    "text": text_cleaned,
                    "start": round(float(snippet.start), 2),
                    "duration": round(float(snippet.duration), 2),
                })

        full_transcript = " ".join(s["text"] for s in segments)

        return {
            "video_id": video_id,
            "title": title,
            "transcript": full_transcript,
            "segments": segments,
        }

    except TranscriptsDisabled as err:
        raise TranscriptsDisabledError(
            f"Transcripts are disabled for YouTube video '{video_id}'. "
            "Subtitles or closed captions must be enabled by the video author."
        ) from err

    except NoTranscriptFound as err:
        raise TranscriptNotFoundError(
            f"No transcript found for YouTube video '{video_id}'. "
            "No subtitles are available in English or any fallback language."
        ) from err

    except (VideoUnavailable, VideoUnplayable) as err:
        raise VideoUnavailableError(
            f"YouTube video '{video_id}' is unavailable. "
            "It may be private, age-restricted, removed, or region-locked."
        ) from err

    except InvalidVideoId as err:
        raise ValueError(
            f"Invalid YouTube video ID '{video_id}'. Please check the URL."
        ) from err

    except (YouTubeParserError, ValueError):
        raise

    except YouTubeTranscriptApiException as err:
        raise YouTubeParserError(
            f"YouTube transcript extraction failed for video '{video_id}': {err}"
        ) from err

    except Exception as err:
        raise YouTubeParserError(
            f"Unexpected error while parsing YouTube video '{video_id}': {err}"
        ) from err
