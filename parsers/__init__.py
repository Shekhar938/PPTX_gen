"""
parsers package - Transcript and document parsers for PPTX generation.

Provides:
- YouTube transcript parsing (extract_youtube_transcript, extract_video_id)
- Multi-format document parsing (parse_pdf, parse_docx, parse_pptx, parse_document)
"""

from parsers.youtube_parser import (
    extract_youtube_transcript,
    extract_video_id,
    YouTubeParserError,
    TranscriptsDisabledError,
    TranscriptNotFoundError,
    VideoUnavailableError,
)
from parsers.document_parser import (
    parse_pdf,
    parse_docx,
    parse_pptx,
    parse_document,
    DocumentParserError,
    UnsupportedFileTypeError,
)

__all__ = [
    "extract_youtube_transcript",
    "extract_video_id",
    "YouTubeParserError",
    "TranscriptsDisabledError",
    "TranscriptNotFoundError",
    "VideoUnavailableError",
    "parse_pdf",
    "parse_docx",
    "parse_pptx",
    "parse_document",
    "DocumentParserError",
    "UnsupportedFileTypeError",
]
