# xgen_doc2chunk/core/functions/table_cell_image.py
"""
Table Cell Image Support - Shared Helpers

Some documents put an image (typically a screenshot of text) inside a table
cell instead of typing the text. Every format-specific extractor used to drop
those images, leaving the cell blank with no image tag anywhere - so OCR never
saw them.

This module holds the format-agnostic pieces used by the PDF / DOCX / HWPX
extractors to place an image tag *at the correct cell*:

- ``CellImageConfig``      : thresholds tuned for cell-sized images
- ``merge_cell_image_tag`` : consistent way to put a tag into cell content
- ``match_images_to_cells``: geometry matching for coordinate-based formats

Formats whose markup already says which cell owns an image (DOCX, HWPX) only
need ``merge_cell_image_tag``. Coordinate-based formats (PDF) also use
``match_images_to_cells``.
"""
import logging
from dataclasses import dataclass
from typing import Dict, Hashable, List, Optional, Sequence, Tuple

logger = logging.getLogger("document-processor")

BBox = Tuple[float, float, float, float]


# ============================================================================
# Configuration
# ============================================================================

@dataclass
class CellImageConfig:
    """Thresholds for treating an image as the content of a table cell.

    The document-level image filters (``min_image_size=50``,
    ``min_image_area=2500``) are deliberately not reused here: a cell-sized
    screenshot is often only a few dozen pixels tall and would be discarded,
    which is one of the reasons these images went missing.

    Attributes:
        min_width: Minimum pixel width for a cell image
        min_height: Minimum pixel height for a cell image
        min_area: Minimum pixel area for a cell image
        min_overlap: Fraction of the image that must sit inside the cell
        min_cell_coverage: Fraction of the *cell* the image must cover.
            This is what separates content from decoration: a screenshot
            pasted as cell content fills most of its cell, while a checkbox
            or bullet icon occupies a few percent of it. Pixel size alone
            cannot tell them apart, because a small screenshot of text and a
            small icon have comparable pixel counts.
        only_empty_cells: Only attach to cells that have no text.
            Keeping this True means tables that already extract correctly
            follow exactly the same code path as before.
    """
    min_width: int = 8
    min_height: int = 8
    min_area: int = 64
    min_overlap: float = 0.6
    min_cell_coverage: float = 0.25
    only_empty_cells: bool = True

    def accepts_size(self, width: int, height: int) -> bool:
        """Check whether an image is large enough to be meaningful content."""
        if width < self.min_width or height < self.min_height:
            return False
        if width * height < self.min_area:
            return False
        return True

    def accepts_coverage(self, coverage: float) -> bool:
        """Check whether the image covers enough of the cell to be content."""
        return coverage >= self.min_cell_coverage


DEFAULT_CELL_IMAGE_CONFIG = CellImageConfig()


# ============================================================================
# Cell content merging
# ============================================================================

def merge_cell_image_tag(content: Optional[str], tag: str) -> str:
    """Place an image tag into a cell's content.

    Existing text is kept and the tag appended, so a cell holding both a
    caption and an image keeps both. Duplicate tags are not appended twice.

    Args:
        content: Current cell content (may be None or empty)
        tag: Image tag such as ``[Image:temp/images/x.png]``

    Returns:
        Updated cell content
    """
    tag = (tag or "").strip()
    if not tag:
        return content or ""

    existing = (content or "").strip()
    if not existing:
        return tag
    if tag in existing:
        return existing
    return f"{existing} {tag}"


# ============================================================================
# Geometry matching (coordinate-based formats)
# ============================================================================

def overlap_ratio(inner: BBox, outer: BBox) -> float:
    """Fraction of *inner*'s area that lies within *outer*."""
    ix0 = max(inner[0], outer[0])
    iy0 = max(inner[1], outer[1])
    ix1 = min(inner[2], outer[2])
    iy1 = min(inner[3], outer[3])

    if ix1 <= ix0 or iy1 <= iy0:
        return 0.0

    inner_area = (inner[2] - inner[0]) * (inner[3] - inner[1])
    if inner_area <= 0:
        return 0.0

    return ((ix1 - ix0) * (iy1 - iy0)) / inner_area


def match_images_to_cells(
    cells: Sequence[Tuple[Hashable, BBox]],
    images: Sequence[Tuple[Hashable, BBox]],
    min_overlap: float = 0.6,
) -> Dict[Hashable, List[Hashable]]:
    """Assign each image to the single cell that contains most of it.

    An image is assigned to at most one cell, so an image spanning a border
    cannot be duplicated into both neighbours.

    Args:
        cells: (cell_key, bbox) pairs
        images: (image_key, bbox) pairs
        min_overlap: Minimum fraction of the image inside the cell

    Returns:
        cell_key -> list of image_keys, in the order given by *images*
    """
    result: Dict[Hashable, List[Hashable]] = {}

    if not cells or not images:
        return result

    for image_key, image_bbox in images:
        best_key = None
        best_ratio = 0.0

        for cell_key, cell_bbox in cells:
            ratio = overlap_ratio(image_bbox, cell_bbox)
            if ratio > best_ratio:
                best_key = cell_key
                best_ratio = ratio

        if best_key is not None and best_ratio >= min_overlap:
            result.setdefault(best_key, []).append(image_key)

    return result


__all__ = [
    "BBox",
    "CellImageConfig",
    "DEFAULT_CELL_IMAGE_CONFIG",
    "merge_cell_image_tag",
    "overlap_ratio",
    "match_images_to_cells",
]
