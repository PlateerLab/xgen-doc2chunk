# xgen_doc2chunk/core/processor/pdf_helpers/pdf_cell_image.py
"""
PDF Table Cell Image Attachment

When a PDF table cell holds an image instead of text (a pasted screenshot of
text, a stamp, a rendered formula), the cell extracts as empty. The image is
also skipped by the document-level image pass, which drops anything sitting
inside a table bbox (``pdf_image_processor.extract_images_from_page``), so no
``[Image:...]`` tag is produced anywhere and OCR never sees it.

This module fills that gap: it matches image placements to the table cell that
contains them and writes an image tag into that cell's data slot, so the tag
ends up inside the corresponding ``<td>`` of the generated table HTML and is
picked up by the normal OCR tag-replacement pass.

Design notes
------------
- Only cells that extracted as **empty** are touched (``only_empty_cells``).
  Tables that already extract correctly take the exact same path as before.
- The document-level image pass is left untouched. Its table-bbox exclusion
  still prevents the same placement being emitted twice.
- Failure at any point leaves the cell empty, i.e. current behaviour.
"""
import logging
from typing import Any, Dict, List, Optional, Set, Tuple

from xgen_doc2chunk.core.functions.table_cell_image import (
    CellImageConfig,
    DEFAULT_CELL_IMAGE_CONFIG,
    match_images_to_cells,
    merge_cell_image_tag,
    overlap_ratio,
)

logger = logging.getLogger("document-processor")


def collect_page_image_placements(page) -> List[Dict[str, Any]]:
    """List every image placement on a page with its position.

    ``get_image_info(xrefs=True)`` reports one entry per *placement*, so an
    image reused in several cells is reported once per location.

    Args:
        page: PyMuPDF page object

    Returns:
        List of dicts with 'xref', 'bbox', 'width', 'height'
    """
    placements: List[Dict[str, Any]] = []

    try:
        infos = page.get_image_info(xrefs=True)
    except Exception as e:
        logger.debug(f"[PDF] get_image_info failed on page {page.number}: {e}")
        return placements

    for info in infos:
        xref = info.get("xref", 0)
        bbox = info.get("bbox")
        if not bbox:
            continue
        placements.append({
            "xref": xref,
            "bbox": tuple(bbox),
            "width": info.get("width", 0),
            "height": info.get("height", 0),
        })

    return placements


def _cell_is_empty(data: List[List[Optional[str]]], row: int, col: int) -> bool:
    """True when the cell has no extracted text."""
    if row < 0 or row >= len(data):
        return False
    row_data = data[row]
    if row_data is None or col < 0 or col >= len(row_data):
        return False
    return not str(row_data[col] or "").strip()


def _load_image_bytes(doc, xref: int) -> Tuple[Optional[bytes], int, int]:
    """Fetch decoded image bytes and pixel dimensions for an xref."""
    try:
        base_image = doc.extract_image(xref)
    except Exception as e:
        logger.debug(f"[PDF] extract_image failed for xref={xref}: {e}")
        return None, 0, 0

    if not base_image:
        return None, 0, 0

    return (
        base_image.get("image"),
        base_image.get("width", 0),
        base_image.get("height", 0),
    )


def attach_cell_images_to_candidate(
    page,
    candidate,
    image_processor,
    cell_image_xrefs: Optional[Set[int]] = None,
    config: Optional[CellImageConfig] = None,
) -> int:
    """Write image tags into the table cells that contain images.

    Args:
        page: PyMuPDF page the table was detected on
        candidate: TableCandidate with ``.data`` and ``.cells`` already final
            (narrow-column merging and header/data merging applied)
        image_processor: Format image processor exposing ``process_image``
        cell_image_xrefs: Set updated with xrefs consumed as cell content.
            Kept separate from the document-level ``processed_images`` set so
            the existing image pass is not affected.
        config: Cell image thresholds

    Returns:
        Number of cells that received an image tag
    """
    if image_processor is None or candidate is None:
        return 0
    if not getattr(candidate, "data", None) or not getattr(candidate, "cells", None):
        return 0

    cfg = config or DEFAULT_CELL_IMAGE_CONFIG

    # Cells that can receive an image, keyed by (row, col)
    cell_entries: List[Tuple[Tuple[int, int], Tuple[float, float, float, float]]] = []
    for cell in candidate.cells:
        bbox = getattr(cell, "bbox", None)
        if not bbox or len(bbox) < 4:
            continue
        # A zero-area bbox carries no position information
        if bbox[2] <= bbox[0] or bbox[3] <= bbox[1]:
            continue
        if cfg.only_empty_cells and not _cell_is_empty(candidate.data, cell.row, cell.col):
            continue
        cell_entries.append(((cell.row, cell.col), tuple(bbox[:4])))

    if not cell_entries:
        return 0

    placements = collect_page_image_placements(page)
    if not placements:
        return 0

    image_entries = [(idx, p["bbox"]) for idx, p in enumerate(placements)]
    matches = match_images_to_cells(cell_entries, image_entries, cfg.min_overlap)
    if not matches:
        return 0

    cell_bboxes = dict(cell_entries)
    doc = page.parent
    page_num = page.number
    attached = 0

    for (row, col), image_indices in matches.items():
        tags: List[str] = []
        cell_bbox = cell_bboxes.get((row, col))

        for image_idx in image_indices:
            placement = placements[image_idx]
            xref = placement["xref"]

            # Content fills its cell; a checkbox or bullet icon does not.
            if cell_bbox is not None:
                coverage = overlap_ratio(cell_bbox, placement["bbox"])
                if not cfg.accepts_coverage(coverage):
                    logger.debug(
                        f"[PDF] Cell image covers only {coverage:.0%} of the cell, "
                        f"treated as decoration: xref={xref}"
                    )
                    continue

            image_bytes, width, height = _load_image_bytes(doc, xref)
            if not image_bytes:
                continue

            # Cell-sized images are small by nature; the document-level
            # thresholds would discard them.
            if not cfg.accepts_size(width, height):
                logger.debug(
                    f"[PDF] Cell image too small, skipped: xref={xref} {width}x{height}"
                )
                continue

            try:
                tag = image_processor.process_image(
                    image_bytes, xref=xref, page_num=page_num
                )
            except Exception as e:
                logger.warning(f"[PDF] Cell image save failed xref={xref}: {e}")
                continue

            if not tag:
                continue

            tags.append(tag)
            if cell_image_xrefs is not None:
                cell_image_xrefs.add(xref)

        if not tags:
            continue

        # Guard against cell coordinates outside the data grid
        if row >= len(candidate.data):
            continue
        row_data = candidate.data[row]
        if row_data is None or col >= len(row_data):
            continue

        content = row_data[col]
        for tag in tags:
            content = merge_cell_image_tag(content, tag)
        row_data[col] = content
        attached += 1

    if attached:
        logger.info(
            f"[PDF] Page {page_num + 1}: attached image tags to {attached} table cell(s)"
        )

    return attached


__all__ = [
    "collect_page_image_placements",
    "attach_cell_images_to_candidate",
]
