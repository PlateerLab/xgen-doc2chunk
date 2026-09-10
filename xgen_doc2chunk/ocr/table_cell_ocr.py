# xgen_doc2chunk/ocr/table_cell_ocr.py
"""
Table-Cell Aware OCR Tag Replacement

Image tags may appear anywhere in extracted text, including *inside* an HTML
table cell (``<td>[Image:...]</td>``) when a document has an image embedded in
a table cell.

Replacing a tag that sits inside a cell needs two things that plain
``re.sub`` does not give:

1. **Literal replacement.**
   ``pattern.sub(replacement_string, text)`` interprets backslashes and group
   references (``\\1``, ``\\g<name>``) inside *replacement_string*. VL model
   output routinely contains backslashes, which either corrupts the result or
   raises ``re.error``. Using a replacement *function* makes the value literal.

2. **Cell-safe output.**
   The default OCR prompt asks the VL model to return HTML tables. Dropping
   ``<table><tr><td>...`` inside an existing ``<td>`` breaks the downstream
   table parser (``chunking/table_parser.py`` matches cells with
   ``<td[^>]*>(.*?)</td>``). Text landing inside a cell is therefore flattened
   to plain inline text.

Text *outside* table cells is left exactly as before - no behavioural change
for documents without images in tables.
"""
import bisect
import logging
import re
from typing import Callable, Optional, Pattern

logger = logging.getLogger("ocr-processor")


# ============================================================================
# Table cell boundary detection
# ============================================================================

_CELL_OPEN_RE = re.compile(r'<(?:td|th)\b[^>]*>', re.IGNORECASE)
_CELL_CLOSE_RE = re.compile(r'</(?:td|th)\s*>', re.IGNORECASE)


class TableCellIndex:
    """Answers "is this offset inside an HTML table cell?" in O(log n).

    Boundaries are indexed once per document instead of rescanning the text
    for every image tag.
    """

    def __init__(self, text: str):
        events = [(m.end(), True) for m in _CELL_OPEN_RE.finditer(text)]
        events += [(m.end(), False) for m in _CELL_CLOSE_RE.finditer(text)]
        events.sort(key=lambda e: e[0])
        self._positions = [e[0] for e in events]
        self._is_open = [e[1] for e in events]

    def is_inside(self, pos: int) -> bool:
        """True if *pos* falls between a cell open tag and its close tag."""
        if not self._positions:
            return False
        idx = bisect.bisect_right(self._positions, pos) - 1
        if idx < 0:
            return False
        return self._is_open[idx]


# ============================================================================
# Cell-safe text flattening
# ============================================================================

_FENCE_OPEN_RE = re.compile(r'^\s*```[A-Za-z0-9_+-]*\s*')
_FENCE_CLOSE_RE = re.compile(r'\s*```\s*$')
_BR_RE = re.compile(r'<br\s*/?>', re.IGNORECASE)
_TABLE_STRUCT_RE = re.compile(
    r'</?(?:table|thead|tbody|tfoot|tr|td|th|caption|colgroup|col)\b[^>]*>',
    re.IGNORECASE,
)
_ANY_TAG_RE = re.compile(r'<[^>]*>')
_MD_SEPARATOR_RE = re.compile(r'(?m)^[ \t]*\|?[ \t]*:?-{2,}[-\s:|]*$')
_WS_RE = re.compile(r'\s+')


def sanitize_ocr_text_for_table_cell(text: str) -> str:
    """Flatten OCR output so it can live inside an HTML ``<td>``.

    Removes markdown code fences, converts table/line markup to spaces, strips
    any remaining tags, collapses whitespace, then escapes HTML special
    characters so the surrounding table markup stays parseable.

    Args:
        text: Raw OCR/VL model output

    Returns:
        Single-line, HTML-escaped text safe to place inside a table cell
    """
    if not text:
        return ""

    s = text.strip()

    # ``` fenced blocks -> bare content
    s = _FENCE_OPEN_RE.sub('', s)
    s = _FENCE_CLOSE_RE.sub('', s)

    # Line breaks and table structure become separators, not markup
    s = _BR_RE.sub(' ', s)
    s = _TABLE_STRUCT_RE.sub(' ', s)

    # Anything else that still looks like a tag is dropped
    s = _ANY_TAG_RE.sub(' ', s)

    # Markdown table leftovers
    s = _MD_SEPARATOR_RE.sub(' ', s)
    s = s.replace('|', ' ')

    s = _WS_RE.sub(' ', s).strip()

    # Escape last so the cell markup around it cannot be broken.
    # All real tags were removed above, so remaining '<' '>' '&' are literal.
    s = s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

    return s


# ============================================================================
# Replacement driver
# ============================================================================

def replace_image_tags(
    text: str,
    pattern: Pattern[str],
    resolve: Callable[[str], Optional[str]],
) -> str:
    """Replace every image tag with its OCR result, cell-aware and escape-safe.

    ``resolve`` is called at most once per distinct image path; its result is
    reused for repeated tags pointing at the same file. Returning ``None``
    keeps the original tag untouched (load/convert failure).

    Args:
        text: Text containing image tags
        pattern: Compiled tag pattern whose group 1 is the image path
        resolve: path -> OCR text, or None to keep the tag

    Returns:
        Text with tags replaced
    """
    if not text:
        return text

    cell_index = TableCellIndex(text)
    cache = {}

    def _replace(match: 're.Match') -> str:
        path = match.group(1)

        if path not in cache:
            try:
                cache[path] = resolve(path)
            except Exception as exc:  # a single bad image must not kill the doc
                logger.warning(f"[OCR] Tag resolution failed for {path}: {exc}")
                cache[path] = None

        result = cache[path]
        if result is None:
            return match.group(0)

        if cell_index.is_inside(match.start()):
            flattened = sanitize_ocr_text_for_table_cell(result)
            if not flattened:
                return match.group(0)
            return flattened

        return result

    return pattern.sub(_replace, text)


__all__ = [
    "TableCellIndex",
    "sanitize_ocr_text_for_table_cell",
    "replace_image_tags",
]
