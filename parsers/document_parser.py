"""
document_parser.py - Unified Document Parser for PPTX Generation.

Extracts text, metadata, structural sections, tables, and slides from:
- PDF files (using pymupdf / fitz)
- Word documents (using python-docx)
- PowerPoint decks (using python-pptx)
Provides a unified entry point `parse_document()` that auto-detects format.
"""

from __future__ import annotations

import logging
import re
import warnings
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

# Suppress pymupdf import deprecation warning for fitz
warnings.filterwarnings("ignore", category=FutureWarning, module="fitz")

try:
    import fitz  # pymupdf
except ImportError:
    import pymupdf as fitz

import docx
from docx.table import Table as DocxTable
from docx.text.paragraph import Paragraph as DocxParagraph
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

from utils.helpers import sanitize_text

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Custom Exceptions
# ---------------------------------------------------------------------------
class DocumentParserError(Exception):
    """Base exception for document parsing errors."""
    pass


class UnsupportedFileTypeError(DocumentParserError):
    """Raised when an unsupported file extension is provided."""
    pass


# ---------------------------------------------------------------------------
# PDF Parser
# ---------------------------------------------------------------------------
def parse_pdf(file_path: str) -> Dict[str, Any]:
    """Extract text, tables, and metadata from a PDF file using PyMuPDF (fitz).

    Extracts textual content page-by-page, detects tabular data using PyMuPDF's
    table finder, and parses standard document metadata (title, author, creation date,
    and table of contents if available).

    Args:
        file_path: Path to the PDF document.

    Returns:
        Dict[str, Any]: Extracted content with structure:
            - 'text' (str): Full concatenated text across all pages.
            - 'pages' (int): Total number of pages.
            - 'metadata' (dict): Metadata dictionary (title, author, subject, keywords,
              creator, producer, creation_date, mod_date, toc).
            - 'sections' (list[dict]): Per-page list of dictionaries containing:
              'page_number' (int), 'heading' (str), 'content' (str), 'tables' (list).

    Raises:
        FileNotFoundError: If the file does not exist on disk.
        ValueError: If the file is not a valid or readable PDF.
    """
    path = Path(file_path).resolve()
    if not path.is_file():
        raise FileNotFoundError(f"PDF file not found at path: {file_path}")

    try:
        doc = fitz.open(str(path))
    except Exception as err:
        raise ValueError(f"Failed to open or parse PDF file '{path.name}': {err}") from err

    try:
        page_count = doc.page_count
        raw_meta = doc.metadata or {}
        toc = doc.get_toc() or []

        metadata: Dict[str, Any] = {
            "title": sanitize_text(raw_meta.get("title") or ""),
            "author": sanitize_text(raw_meta.get("author") or ""),
            "subject": sanitize_text(raw_meta.get("subject") or ""),
            "keywords": sanitize_text(raw_meta.get("keywords") or ""),
            "creator": sanitize_text(raw_meta.get("creator") or ""),
            "producer": sanitize_text(raw_meta.get("producer") or ""),
            "creation_date": raw_meta.get("creationDate") or "",
            "mod_date": raw_meta.get("modDate") or "",
            "toc": toc,
        }

        sections: List[Dict[str, Any]] = []
        page_texts: List[str] = []

        for page_idx in range(page_count):
            page_num = page_idx + 1
            page = doc.load_page(page_idx)

            # Extract raw text from page
            raw_text = page.get_text()
            cleaned_text = sanitize_text(raw_text)

            # Extract tables if present on the page
            page_tables: List[List[List[str]]] = []
            try:
                table_finder = page.find_tables()
                for table in table_finder:
                    extracted = table.extract()
                    if extracted:
                        clean_table_rows = [
                            [sanitize_text(str(cell or "")) for cell in row]
                            for row in extracted
                        ]
                        page_tables.append(clean_table_rows)
            except Exception as table_err:
                logger.debug(f"Could not extract tables on page {page_num} of {path.name}: {table_err}")

            # Append formatted table content to page text if tables exist
            if page_tables:
                table_text_blocks: List[str] = []
                for t_idx, tbl in enumerate(page_tables, start=1):
                    row_strs = [" | ".join(row) for row in tbl]
                    table_text_blocks.append(f"[Table {t_idx}]\n" + "\n".join(row_strs))
                if table_text_blocks:
                    table_combined = "\n\n".join(table_text_blocks)
                    if cleaned_text:
                        cleaned_text = f"{cleaned_text}\n\n{table_combined}"
                    else:
                        cleaned_text = table_combined

            sections.append({
                "page_number": page_num,
                "heading": f"Page {page_num}",
                "content": cleaned_text,
                "tables": page_tables,
            })

            if cleaned_text:
                page_texts.append(cleaned_text)

        full_text = "\n\n".join(page_texts)

        return {
            "text": full_text,
            "pages": page_count,
            "metadata": metadata,
            "sections": sections,
        }

    finally:
        doc.close()


# ---------------------------------------------------------------------------
# DOCX Parser
# ---------------------------------------------------------------------------
def _extract_heading_level(style_name: str) -> Optional[int]:
    """Helper to detect heading level from style name."""
    if not style_name:
        return None

    style_clean = style_name.strip()
    match = re.search(r"Heading\s*(\d+)", style_clean, re.IGNORECASE)
    if match:
        return int(match.group(1))

    if style_clean.lower() == "title":
        return 1
    if style_clean.lower() == "subtitle":
        return 2

    return None


def parse_docx(file_path: str) -> Dict[str, Any]:
    """Extract text from a Word (.docx) file while preserving heading structure.

    Iterates through document paragraphs and tables in order, grouping body content
    under headings and preserving hierarchical levels.

    Args:
        file_path: Path to the .docx document.

    Returns:
        Dict[str, Any]: Extracted content with structure:
            - 'text' (str): Full text of the document.
            - 'sections' (list[dict]): Hierarchical sections where each item is:
              {'heading': str, 'content': str, 'level': int}.
            - 'metadata' (dict): Document core properties (title, author, subject, etc.).

    Raises:
        FileNotFoundError: If the file does not exist on disk.
        ValueError: If the file is not a valid .docx document.
    """
    path = Path(file_path).resolve()
    if not path.is_file():
        raise FileNotFoundError(f"DOCX file not found at path: {file_path}")

    try:
        doc = docx.Document(str(path))
    except Exception as err:
        raise ValueError(f"Failed to open or parse DOCX file '{path.name}': {err}") from err

    # Extract metadata from document core properties
    core_props = doc.core_properties
    metadata: Dict[str, Any] = {
        "title": sanitize_text(core_props.title or ""),
        "author": sanitize_text(core_props.author or ""),
        "subject": sanitize_text(core_props.subject or ""),
        "keywords": sanitize_text(core_props.keywords or ""),
        "category": sanitize_text(core_props.category or ""),
        "created": str(core_props.created) if core_props.created else "",
        "modified": str(core_props.modified) if core_props.modified else "",
    }

    sections: List[Dict[str, Any]] = []
    current_heading = "Introduction"
    current_level = 1
    current_content: List[str] = []
    full_text_blocks: List[str] = []

    def commit_section() -> None:
        nonlocal current_heading, current_level, current_content
        joined_content = sanitize_text("\n\n".join(current_content))
        # Commit if there is content or if heading is not the default empty intro
        if joined_content or (current_heading != "Introduction" and sections):
            sections.append({
                "heading": current_heading,
                "content": joined_content,
                "level": current_level,
            })
        current_content = []

    # Iterate over body elements in true document order
    for child in doc.element.body:
        if child.tag.endswith("p"):
            p = DocxParagraph(child, doc)
            text = sanitize_text(p.text)
            if not text:
                continue

            heading_lvl = _extract_heading_level(p.style.name if p.style else "")
            if heading_lvl is not None:
                # Flush previous section
                commit_section()
                current_heading = text
                current_level = heading_lvl
                full_text_blocks.append(f"\n## {text}\n")
            else:
                current_content.append(text)
                full_text_blocks.append(text)

        elif child.tag.endswith("tbl"):
            t = DocxTable(child, doc)
            table_rows: List[str] = []
            for row in t.rows:
                cells = [sanitize_text(c.text) for c in row.cells]
                # Avoid duplicate identical adjacent cells caused by merged cell spans
                deduped_cells: List[str] = []
                for c in cells:
                    if not deduped_cells or c != deduped_cells[-1]:
                        deduped_cells.append(c)
                table_rows.append(" | ".join(deduped_cells))

            if table_rows:
                table_str = "\n".join(table_rows)
                current_content.append(table_str)
                full_text_blocks.append(table_str)

    # Commit any trailing section
    commit_section()

    # If no heading was present in the whole document, ensure a single section
    if not sections and current_content:
        sections.append({
            "heading": metadata["title"] or "Document Content",
            "content": sanitize_text("\n\n".join(current_content)),
            "level": 1,
        })

    full_text = sanitize_text("\n\n".join(full_text_blocks))

    return {
        "text": full_text,
        "sections": sections,
        "metadata": metadata,
    }


# ---------------------------------------------------------------------------
# PPTX Parser
# ---------------------------------------------------------------------------
def _extract_shape_text(shape: Any) -> List[str]:
    """Helper to recursively extract text from PPTX shapes, tables, and groups."""
    texts: List[str] = []

    # Text frames
    if shape.has_text_frame:
        frame_text = sanitize_text(shape.text_frame.text)
        if frame_text:
            texts.append(frame_text)

    # Tables
    elif shape.has_table:
        table_rows: List[str] = []
        for row in shape.table.rows:
            row_cells = [sanitize_text(cell.text) for cell in row.cells]
            table_rows.append(" | ".join(row_cells))
        if table_rows:
            texts.append("\n".join(table_rows))

    # Group shapes
    elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        for sub_shape in shape.shapes:
            texts.extend(_extract_shape_text(sub_shape))

    return texts


def parse_pptx(file_path: str) -> Dict[str, Any]:
    """Extract slide content and structure from an existing PowerPoint (.pptx) file.

    Extracts titles, text frames, tables, group shapes, and speaker notes
    for each slide.

    Args:
        file_path: Path to the .pptx presentation file.

    Returns:
        Dict[str, Any]: Extracted presentation content with structure:
            - 'text' (str): Full concatenated text across all slides.
            - 'slides' (list[dict]): List of slide dictionaries with keys:
              'slide_number' (int), 'title' (str), 'content' (str).
            - 'slide_count' (int): Total number of slides.

    Raises:
        FileNotFoundError: If the file does not exist on disk.
        ValueError: If the file is not a valid .pptx presentation.
    """
    path = Path(file_path).resolve()
    if not path.is_file():
        raise FileNotFoundError(f"PPTX file not found at path: {file_path}")

    try:
        prs = Presentation(str(path))
    except Exception as err:
        raise ValueError(f"Failed to open or parse PPTX file '{path.name}': {err}") from err

    slides: List[Dict[str, Any]] = []
    full_text_parts: List[str] = []

    for idx, slide in enumerate(prs.slides, start=1):
        # Determine title
        title = ""
        title_shape = None
        if slide.shapes.title and slide.shapes.title.has_text_frame:
            title_shape = slide.shapes.title
            title = sanitize_text(title_shape.text_frame.text)

        # Collect body content from other shapes
        body_parts: List[str] = []
        for shape in slide.shapes:
            if shape == title_shape:
                continue
            shape_texts = _extract_shape_text(shape)
            body_parts.extend(shape_texts)

        # If no explicit title shape was identified, use first line or fallback
        if not title:
            if body_parts:
                first_text = body_parts[0]
                lines = first_text.split("\n")
                title = lines[0].strip()
                # If the first item had multiple lines, keep remainder in body
                if len(lines) > 1:
                    body_parts[0] = "\n".join(lines[1:]).strip()
                else:
                    body_parts.pop(0)
            else:
                title = f"Slide {idx}"

        # Collect speaker notes if present
        if slide.has_notes_slide:
            notes_frame = slide.notes_slide.notes_text_frame
            notes_text = sanitize_text(notes_frame.text) if notes_frame else ""
            if notes_text:
                body_parts.append(f"[Speaker Notes]: {notes_text}")

        slide_content = sanitize_text("\n\n".join(body_parts))
        slides.append({
            "slide_number": idx,
            "title": title,
            "content": slide_content,
        })

        slide_text_block = f"--- Slide {idx}: {title} ---\n{slide_content}".strip()
        full_text_parts.append(slide_text_block)

    full_text = "\n\n".join(full_text_parts)

    return {
        "text": full_text,
        "slides": slides,
        "slide_count": len(slides),
    }


# ---------------------------------------------------------------------------
# Unified Document Parser Entry Point
# ---------------------------------------------------------------------------
def parse_document(file_path: str) -> Dict[str, Any]:
    """Unified entry point that auto-detects file type by extension and parses it.

    Supported formats:
    - PDF (.pdf) -> calls `parse_pdf()`
    - Word (.docx) -> calls `parse_docx()`
    - PowerPoint (.pptx) -> calls `parse_pptx()`

    Returns the parser dictionary with an added `'source_type'` field ('pdf', 'docx', 'pptx').

    Args:
        file_path: Path to the document file.

    Returns:
        Dict[str, Any]: Parsed document dictionary including 'source_type'.

    Raises:
        FileNotFoundError: If the document file does not exist.
        ValueError: If file is corrupted or cannot be parsed.
        UnsupportedFileTypeError: If file extension is unsupported (e.g. .txt, .doc).
    """
    path = Path(file_path).resolve()
    if not path.is_file():
        raise FileNotFoundError(f"Document file not found at path: {file_path}")

    suffix = path.suffix.lower()

    if suffix == ".pdf":
        result = parse_pdf(str(path))
        result["source_type"] = "pdf"
        return result

    elif suffix == ".docx":
        result = parse_docx(str(path))
        result["source_type"] = "docx"
        return result

    elif suffix == ".pptx":
        result = parse_pptx(str(path))
        result["source_type"] = "pptx"
        return result

    elif suffix in (".doc", ".ppt"):
        raise UnsupportedFileTypeError(
            f"Legacy binary format '{suffix}' is not supported directly for '{path.name}'. "
            f"Please convert the file to modern '{suffix}x' format."
        )

    else:
        raise UnsupportedFileTypeError(
            f"Unsupported document format '{suffix}' for file '{path.name}'. "
            "Supported formats are: .pdf, .docx, .pptx"
        )
