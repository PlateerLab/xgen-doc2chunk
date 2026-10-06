# xgen_doc2chunk/chunking/chunking.py
"""
Document Chunking Module - Advanced Text Chunking System

Main Features:
- HTML table-preserving chunking with row-level splitting
- Markdown table-preserving chunking with row-level splitting
- Intelligent splitting for large table data (CSV/TSV/Excel)
- Table structure restoration (header preservation for both HTML and Markdown)
- Page-based chunking
- Language-specific code file chunking

Key Improvements (Table Chunking Enhancement):
- Split large tables (HTML and Markdown) by rows to fit chunk_size
- Automatically restore table headers in each chunk
- Ensure table structure integrity
- Add chunk indexing metadata
- NO OVERLAP for table chunks (intentional to prevent data duplication)

Critical Rules for Table-Based Files (CSV, TSV, XLSX, XLS):
- Always use force_chunking=True
- Always split by rows (never cut in the middle of a row)
- Never apply overlap between table chunks
- Restore headers in each chunk for context

Refactoring:
- Core logic is separated into chunking_helper submodules
- This file maintains only the public API and integration logic
"""
import bisect
import functools
import logging
import re
from typing import Any, Dict, List, Optional, Union

# Import from individual modules
from xgen_doc2chunk.chunking.constants import (
    TABLE_SIZE_THRESHOLD_MULTIPLIER,
    TABLE_BASED_FILE_TYPES,
    HTML_TABLE_PATTERN,
    MARKDOWN_TABLE_PATTERN,
)
from xgen_doc2chunk.chunking.table_chunker import (
    chunk_large_table as _chunk_large_table,
    chunk_large_markdown_table as _chunk_large_markdown_table,
    is_markdown_table as _is_markdown_table,
)

from xgen_doc2chunk.chunking.protected_regions import (
    find_protected_regions as _find_protected_regions,
    get_protected_region_positions as _get_protected_region_positions,
    split_with_protected_regions as _split_with_protected_regions,
)

from xgen_doc2chunk.chunking.page_chunker import (
    chunk_by_pages as _chunk_by_pages,
)

from xgen_doc2chunk.chunking.text_chunker import (
    chunk_plain_text as _chunk_plain_text,
    chunk_text_without_tables,
    chunk_with_row_protection,
    clean_chunks as _clean_chunks,
    reconstruct_text_from_chunks,
    find_overlap_length,
)

from xgen_doc2chunk.chunking.sheet_processor import (
    extract_document_metadata as _extract_document_metadata,
    prepend_metadata_to_chunks as _prepend_metadata_to_chunks,
    extract_sheet_sections as _extract_sheet_sections,
    chunk_multi_sheet_content,
    chunk_single_table_content,
)

logger = logging.getLogger("document-processor")


# ============================================================================
# Helper Functions for PageTagProcessor integration
# ============================================================================

def _get_page_marker_patterns(page_tag_processor: Optional[Any] = None) -> List[str]:
    """
    Get page marker regex patterns from PageTagProcessor or use defaults.
    
    Args:
        page_tag_processor: PageTagProcessor instance (optional)
        
    Returns:
        List of regex patterns for page/slide markers
    """
    if page_tag_processor is not None:
        # Build patterns from processor's config
        config = page_tag_processor.config
        patterns = [
            page_tag_processor.get_pattern_string(),  # Page pattern
        ]
        # Add slide pattern if different prefix
        if config.slide_prefix != config.tag_prefix:
            from xgen_doc2chunk.core.functions.page_tag_processor import PageTagType
            patterns.append(page_tag_processor.get_pattern_string(PageTagType.SLIDE))
        return patterns
    else:
        # Default patterns
        return [
            r'\[Page Number:\s*(\d+)\]',
            r'\[Slide Number:\s*(\d+)\]',
        ]


def _get_sheet_marker_pattern(page_tag_processor: Optional[Any] = None) -> str:
    """
    Get sheet marker regex pattern from PageTagProcessor or use default.
    
    Args:
        page_tag_processor: PageTagProcessor instance (optional)
        
    Returns:
        Regex pattern for sheet markers
    """
    if page_tag_processor is not None:
        from xgen_doc2chunk.core.functions.page_tag_processor import PageTagType
        return page_tag_processor.get_pattern_string(PageTagType.SHEET)
    else:
        return r'\[Sheet:\s*([^\]]+)\]'


def _get_image_tag_pattern(image_processor: Optional[Any] = None) -> str:
    """
    Get image tag regex pattern from ImageProcessor or use default.
    
    Args:
        image_processor: ImageProcessor instance (optional)
        
    Returns:
        Regex pattern for image tags
    """
    if image_processor is not None:
        return image_processor.get_pattern_string()
    else:
        # Default pattern: [Image:...] or [image:...] with optional spaces and braces
        from xgen_doc2chunk.chunking.constants import IMAGE_TAG_PATTERN
        return IMAGE_TAG_PATTERN


def _get_chart_block_pattern(chart_processor: Optional[Any] = None) -> str:
    """
    Get chart block regex pattern from ChartProcessor or use default.
    
    Args:
        chart_processor: ChartProcessor instance (optional)
        
    Returns:
        Regex pattern for chart blocks
    """
    if chart_processor is not None:
        try:
            # Build pattern from processor's config
            prefix = re.escape(chart_processor.config.tag_prefix)
            suffix = re.escape(chart_processor.config.tag_suffix)
            return f'{prefix}.*?{suffix}'
        except Exception:
            pass
    # Default pattern: [chart]...[/chart]
    from xgen_doc2chunk.chunking.constants import CHART_BLOCK_PATTERN
    return CHART_BLOCK_PATTERN


def _get_metadata_block_pattern(metadata_formatter: Optional[Any] = None) -> str:
    """
    Get metadata block regex pattern from MetadataFormatter or use default.
    
    Args:
        metadata_formatter: MetadataFormatter instance (optional)
        
    Returns:
        Regex pattern for metadata blocks
    """
    if metadata_formatter is not None:
        try:
            # Build pattern from formatter's config
            prefix = re.escape(metadata_formatter.metadata_tag_prefix)
            suffix = re.escape(metadata_formatter.metadata_tag_suffix)
            return f'{prefix}.*?{suffix}'
        except Exception:
            pass
    # Default pattern: <Document-Metadata>...</Document-Metadata>
    from xgen_doc2chunk.chunking.constants import METADATA_BLOCK_PATTERN
    return METADATA_BLOCK_PATTERN


# ============================================================================
# Public API - Single entry point for external use
# ============================================================================

def create_chunks(
    text: str,
    file_extension: str,
    chunk_size: int = 1000,
    chunk_overlap: int = 200,
    force_chunking: bool = False,
    include_position_metadata: bool = True,
    chunking_strategy: str = "recursive",
    page_tag_processor: Optional[Any] = None,
    image_processor: Optional[Any] = None,
    chart_processor: Optional[Any] = None,
    metadata_formatter: Optional[Any] = None,
    stride: Optional[int] = None,
    parent_chunk_size: Optional[int] = None,
    child_chunk_size: Optional[int] = None,
    **kwargs
) -> Union[List[str], List[Dict[str, Any]]]:
    """
    Split text into chunks. (Single public API)

    Args:
        text: Original text
        file_extension: File extension
        chunk_size: Maximum chunk size
        chunk_overlap: Overlap size between chunks (NOT applied to protected regions)
        force_chunking: Force chunking (disable table protection)
        include_position_metadata: Whether to include position metadata
            - True: Include metadata like page_number, line_start, line_end (List[Dict])
            - False: Return only chunk text (List[str])
        chunking_strategy: Chunking strategy (recursive, sliding, hierarchical) - future implementation
        page_tag_processor: PageTagProcessor instance for custom tag patterns
            - If None, uses default patterns [Page Number: n], [Slide Number: n], [Sheet: name]
            - If provided, uses the processor's configured patterns
            - Page/Slide/Sheet tags are protected and NEVER overlap
        image_processor: ImageProcessor instance for custom image tag patterns
            - If None, uses default pattern [Image:...]
            - If provided, uses the processor's configured patterns
            - Image tags are protected and NEVER overlap
        chart_processor: ChartProcessor instance for custom chart tag patterns
            - If None, uses default pattern [chart]...[/chart]
            - If provided, uses the processor's configured patterns
            - Chart blocks are protected and NEVER overlap
        metadata_formatter: MetadataFormatter instance for custom metadata tag patterns
            - If None, uses default pattern <Document-Metadata>...</Document-Metadata>
            - If provided, uses the formatter's configured patterns
            - Metadata blocks are protected and NEVER overlap
        stride: Stride for sliding window strategy - future implementation
        parent_chunk_size: Parent chunk size for hierarchical strategy - future implementation
        child_chunk_size: Child chunk size for hierarchical strategy - future implementation

    Returns:
        When include_position_metadata=True:
            List of chunks with metadata [{"text", "page_number", "line_start", ...}, ...]
        When include_position_metadata=False:
            List of chunk texts ["chunk1", "chunk2", ...]
            
    Protected Regions (NEVER split or overlap):
        - Image tags: [Image:...] or custom pattern
        - Page/Slide/Sheet tags: [Page Number: n], [Slide Number: n], [Sheet: name] or custom
        - Chart blocks: [chart]...[/chart] or custom
        - Metadata blocks: <Document-Metadata>...</Document-Metadata> or custom
        - Tables: Split by rows, each chunk has NO overlap
    """
    # TODO: Implement various chunking strategies based on chunking_strategy
    if chunking_strategy != "recursive":
        logger.warning(
            f"Chunking strategy '{chunking_strategy}' is not yet implemented, "
            "falling back to 'recursive'"
        )

    # Split text into chunks
    chunks = _split_text(
        text, chunk_size, chunk_overlap,
        file_extension=file_extension,
        force_chunking=force_chunking,
        page_tag_processor=page_tag_processor,
        image_processor=image_processor,
        chart_processor=chart_processor,
        metadata_formatter=metadata_formatter
    )

    # Return chunks without metadata
    if not include_position_metadata:
        return chunks

    # Reconstruct text and build line offset table
    reconstructed = reconstruct_text_from_chunks(chunks, chunk_overlap)
    line_table = _build_line_offset_table(reconstructed, file_extension, page_tag_processor)

    # Chunks repeat the document metadata block at their head. The page is taken
    # from the first line after it, otherwise the block (which sits before the
    # chunk's own page/sheet marker) would assign the previous page.
    leading_metadata = re.compile(
        r'\s*(?:' + _get_metadata_block_pattern(metadata_formatter) + r')\s*', re.DOTALL
    )

    # Line starts are computed once; looking them up per chunk made this quadratic
    line_starts = [line["start"] for line in line_table]

    # Add metadata to each chunk
    result: List[Dict[str, Any]] = []
    current_pos = 0

    for idx, chunk in enumerate(chunks):
        start = current_pos
        end = current_pos + len(chunk) - 1

        start_line_idx = _find_line_index_by_pos(start, line_table, line_starts)
        end_line_idx = _find_line_index_by_pos(end, line_table, line_starts)

        line_start = line_table[start_line_idx]["line_num"]
        line_end = line_table[end_line_idx]["line_num"]

        lead = leading_metadata.match(chunk)
        body_start = start + lead.end() if lead and lead.end() < len(chunk) else start
        page_number = line_table[_find_line_index_by_pos(body_start, line_table, line_starts)].get("page", 1)

        result.append({
            "text": chunk,
            "page_number": page_number,
            "line_start": line_start,
            "line_end": line_end,
            "global_start": start,
            "global_end": end,
            "chunk_index": idx
        })

        current_pos += len(chunk)
        if idx < len(chunks) - 1:
            overlap_len = find_overlap_length(chunk, chunks[idx + 1], chunk_overlap)
            current_pos -= overlap_len

    logger.info(f"Created {len(result)} chunks with position metadata")
    return result


# ============================================================================
# Internal Functions - Table-based content processing
# ============================================================================

def _split_table_based_content(
    text: str,
    chunk_size: int,
    chunk_overlap: int,
    page_tag_processor: Optional[Any] = None,
    image_processor: Optional[Any] = None,
    chart_processor: Optional[Any] = None,
    metadata_formatter: Optional[Any] = None,
    force_chunking: bool = False
) -> List[str]:
    """
    Chunk table-based content (CSV/TSV/Excel).

    Split large tables (HTML or Markdown) to fit chunk_size and restore
    table structure in each chunk.

    For multi-sheet Excel files, process each sheet separately.

    CRITICAL: Table chunks have NO overlap to prevent data duplication.
    This is intentional for search/retrieval quality.

    Args:
        text: Full text (metadata + table)
        chunk_size: Maximum chunk size
        chunk_overlap: Not used for tables (kept for API compatibility)
        page_tag_processor: PageTagProcessor for page/sheet tag patterns
        image_processor: ImageProcessor for image tag patterns
        chart_processor: ChartProcessor for chart block patterns
        metadata_formatter: MetadataFormatter for metadata block patterns
        force_chunking: Allow splitting inside oversized rowspan blocks
            (rows tied together by merged cells). False (default) keeps
            the previous behavior where such blocks are never split.

    Returns:
        List of chunks
    """
    if not text or not text.strip():
        return [""]

    # Bind force_chunking so the table chunking policy reaches
    # chunk_large_table through the sheet processor callbacks
    chunk_table_func = functools.partial(_chunk_table_unified, force_chunking=force_chunking)

    # Get metadata pattern from processor
    metadata_pattern = _get_metadata_block_pattern(metadata_formatter)

    # Extract metadata using custom pattern
    metadata_block, text_without_metadata = _extract_document_metadata(text, metadata_pattern)

    # Extract data analysis block (supports both English and Korean tags)
    analysis_pattern = r'(\[Data Analysis\].*?\[/Data Analysis\])\s*'
    analysis_match = re.search(analysis_pattern, text_without_metadata, re.DOTALL)
    analysis_block = ""

    if analysis_match:
        analysis_block = analysis_match.group(1)
        text_without_analysis = (
            text_without_metadata[:analysis_match.start()] +
            text_without_metadata[analysis_match.end():]
        ).strip()
    else:
        text_without_analysis = text_without_metadata

    # Check for multi-sheet (Excel)
    sheets = _extract_sheet_sections(text_without_analysis)

    # Get patterns from processors for protected region detection
    image_pattern = _get_image_tag_pattern(image_processor)
    chart_pattern = _get_chart_block_pattern(chart_processor)
    metadata_pattern = _get_metadata_block_pattern(metadata_formatter)

    if sheets:
        logger.info(f"Multi-sheet Excel detected: {len(sheets)} sheets")
        # Pass 0 for overlap since tables should not have overlap
        return chunk_multi_sheet_content(
            sheets, metadata_block, analysis_block, chunk_size, 0,
            _chunk_plain_text, chunk_table_func,
            image_pattern=image_pattern,
            chart_pattern=chart_pattern,
            metadata_pattern=metadata_pattern
        )

    # Single table/sheet processing
    # Pass 0 for overlap since tables should not have overlap
    return chunk_single_table_content(
        text_without_analysis, metadata_block, analysis_block, chunk_size, 0,
        _chunk_plain_text, chunk_table_func,
        image_pattern=image_pattern,
        chart_pattern=chart_pattern,
        metadata_pattern=metadata_pattern
    )


def _chunk_table_unified(
    table_text: str,
    chunk_size: int,
    chunk_overlap: int,
    context_prefix: str = "",
    force_chunking: bool = False
) -> List[str]:
    """
    Unified table chunking function that handles both HTML and Markdown tables.

    Detects table type and applies appropriate chunking with NO overlap.

    Args:
        table_text: Table content (HTML or Markdown)
        chunk_size: Maximum chunk size
        chunk_overlap: Ignored (tables have no overlap)
        context_prefix: Context to prepend to each chunk
        force_chunking: Allow splitting inside oversized rowspan blocks
            (HTML tables only - Markdown tables have no rowspan)

    Returns:
        List of table chunks
    """
    if _is_markdown_table(table_text):
        return _chunk_large_markdown_table(table_text, chunk_size, 0, context_prefix)
    else:
        return _chunk_large_table(table_text, chunk_size, 0, context_prefix, force_chunking=force_chunking)


def _split_text(
    text: str,
    chunk_size: int,
    chunk_overlap: int,
    file_extension: Optional[str] = None,
    force_chunking: Optional[bool] = False,
    page_tag_processor: Optional[Any] = None,
    image_processor: Optional[Any] = None,
    chart_processor: Optional[Any] = None,
    metadata_formatter: Optional[Any] = None
) -> List[str]:
    """
    Split text into chunks. (Internal use)

    Preserves HTML and Markdown tables with proper row-level chunking.
    Considers page boundaries for chunking.
    Protects all tag regions (image, page, slide, chart, metadata) with NO overlap.

    Core Strategy:
    1. Apply table-based chunking if file_extension is CSV/TSV/Excel (NO overlap for tables)
    2. Apply page-based chunking first if page markers exist
    3. Merge pages based on chunk_size (allow up to 1.5x)
    4. Never cut in the middle of a table or protected tag
    5. Apply overlap ONLY for plain text (NOT for protected regions)

    Protected Regions (NEVER split or overlap):
    - Image tags, Page/Slide/Sheet tags, Chart blocks, Metadata blocks
    - Tables (split by rows with NO overlap)

    Args:
        text: Original text
        chunk_size: Maximum chunk size
        chunk_overlap: Overlap size between chunks (NOT applied to protected regions)
        file_extension: File extension (csv, xlsx, pdf, etc.) - used for table-based processing
        force_chunking: Force chunking (disable table protection except for table-based files)
        page_tag_processor: PageTagProcessor instance for custom tag patterns
        image_processor: ImageProcessor instance for custom image tag patterns
        chart_processor: ChartProcessor instance for custom chart tag patterns
        metadata_formatter: MetadataFormatter instance for custom metadata tag patterns

    Returns:
        List of chunks
    """
    if not text or not text.strip():
        logger.warning("Empty text provided for chunking")
        return [""]

    # === Check for table-based content (CSV/Excel files only) ===
    # Explicitly determine based on file_extension (no text content guessing)
    is_table_based = file_extension and file_extension.lower() in TABLE_BASED_FILE_TYPES

    # Disable table protection if is_table_based or force_chunking is True
    disable_table_protection = is_table_based or force_chunking

    if is_table_based:
        # For table-based files (CSV/Excel), always use table-based chunking
        # This handles both HTML tables and Markdown tables properly
        logger.info(f"Table-based file detected ({file_extension}), using table-based chunking")
        return _split_table_based_content(
            text, chunk_size, chunk_overlap,
            page_tag_processor=page_tag_processor,
            image_processor=image_processor,
            chart_processor=chart_processor,
            metadata_formatter=metadata_formatter,
            force_chunking=force_chunking
        )

    # Get tag patterns from processors or use defaults (needed for metadata extraction)
    metadata_pattern = _get_metadata_block_pattern(metadata_formatter)

    # Extract metadata using custom pattern
    metadata_block, text_without_metadata = _extract_document_metadata(text, metadata_pattern)
    text = text_without_metadata

    # === Check for page markers ===
    # Build patterns from PageTagProcessor or use defaults
    page_marker_patterns = _get_page_marker_patterns(page_tag_processor)
    has_page_markers = any(re.search(pattern, text) for pattern in page_marker_patterns)

    # Get remaining tag patterns from processors or use defaults
    image_pattern = _get_image_tag_pattern(image_processor)
    chart_pattern = _get_chart_block_pattern(chart_processor)

    if has_page_markers:
        # Page-based chunking
        logger.debug("Page markers found, using page-based chunking")
        chunks = _chunk_by_pages(
            text, chunk_size, chunk_overlap, is_table_based, force_chunking,
            page_tag_processor, image_pattern, chart_pattern, metadata_pattern
        )
    else:
        # Find protected regions (HTML tables, chart blocks, Markdown tables, all tags)
        # Disable table protection on force_chunking (other regions are always protected)
        protected_regions = _find_protected_regions(
            text, is_table_based, force_chunking, image_pattern,
            chart_pattern, page_tag_processor, metadata_pattern
        )
        protected_positions = _get_protected_region_positions(protected_regions)

        if protected_positions:
            region_types = set(r[2] for r in protected_regions)
            logger.info(f"Found {len(protected_positions)} protected regions in document: {region_types}")
            chunks = _split_with_protected_regions(
                text, protected_positions, chunk_size, chunk_overlap, force_chunking,
                image_pattern, chart_pattern, page_tag_processor, metadata_pattern
            )
        else:
            # No protected regions: apply row-level chunking if force_chunking
            if disable_table_protection:
                logger.debug("Force chunking enabled, using row-preserving chunking")
                chunks = _chunk_with_row_protection(text, chunk_size, chunk_overlap, force_chunking)
            else:
                logger.debug("No protected blocks found, using standard chunking")
                return _chunk_text_without_tables(text, chunk_size, chunk_overlap, metadata_block, page_tag_processor)

    # Clean chunks
    cleaned_chunks = _clean_chunks(chunks, page_tag_processor)

    # Merge tiny chunks with next chunk (prevents table-title isolation)
    cleaned_chunks = _merge_small_chunks(
        cleaned_chunks, chunk_size, page_tag_processor
    )

    # Add metadata
    cleaned_chunks = _prepend_metadata_to_chunks(cleaned_chunks, metadata_block)

    logger.info(f"Final text split into {len(cleaned_chunks)} chunks")

    return cleaned_chunks

# ============================================================================
# Internal Functions - Small Chunk Merging
# ============================================================================

def _merge_small_chunks(
    chunks: List[str],
    chunk_size: int,
    page_tag_processor: Optional[Any] = None,
    min_ratio: float = 0.1,
) -> List[str]:
    """
    Merge tiny chunks (< min_ratio of chunk_size) with the next chunk.

    Prevents table titles from being isolated in their own chunk when
    split_with_protected_regions cuts right before a table boundary.

    Rules:
    - If chunk size < chunk_size * min_ratio, merge with next chunk
    - Allow chained merging (A+B still small → merge A+B+C)
    - Don't merge across page boundaries ([Page Number:], [Slide Number:], [Sheet:])

    Args:
        chunks: List of chunks
        chunk_size: Target chunk size
        page_tag_processor: PageTagProcessor for page marker detection
        min_ratio: Threshold ratio (default 0.1 = 10%)

    Returns:
        List of merged chunks
    """
    if len(chunks) <= 1:
        return chunks

    threshold = int(chunk_size * min_ratio)

    # Build page/slide/sheet marker patterns
    page_patterns = _get_page_marker_patterns(page_tag_processor)
    sheet_pattern = _get_sheet_marker_pattern(page_tag_processor)

    def _starts_with_page_boundary(text: str) -> bool:
        """Check if text starts with a page/slide/sheet marker."""
        stripped = text.strip()
        for pattern in page_patterns:
            if re.match(pattern, stripped):
                return True
        if re.match(sheet_pattern, stripped):
            return True
        return False

    result = []
    i = 0

    while i < len(chunks):
        current = chunks[i]

        # Chain-merge while current is small
        while len(current) < threshold and i + 1 < len(chunks):
            next_chunk = chunks[i + 1]

            # Don't merge across page boundaries
            if _starts_with_page_boundary(next_chunk):
                break

            current = current + "\n\n" + next_chunk
            i += 1

        # Backward merge: if still small after forward merge attempts
        # (blocked by page boundary), merge into the previous chunk
        if len(current) < threshold and result:
            result[-1] = result[-1] + "\n\n" + current
        else:
            result.append(current)
        i += 1

    if len(result) < len(chunks):
        merged_count = len(chunks) - len(result)
        logger.info(f"Merged {merged_count} small chunks (threshold: {threshold} chars)")

    return result


# ============================================================================
# Internal Wrapper Functions
# ============================================================================

def _chunk_text_without_tables(
    text: str,
    chunk_size: int,
    chunk_overlap: int,
    metadata: Optional[str],
    page_tag_processor: Optional[Any] = None
) -> List[str]:
    """
    Chunking logic for text without tables.
    Wrapper function for chunk_text_without_tables.
    """
    return chunk_text_without_tables(
        text, chunk_size, chunk_overlap, metadata,
        _prepend_metadata_to_chunks,
        page_tag_processor
    )


def _chunk_with_row_protection(
    text: str,
    chunk_size: int,
    chunk_overlap: int,
    force_chunking: bool = False
) -> List[str]:
    """
    Chunk while protecting row boundaries when table protection is disabled.
    
    Both HTML and Markdown tables are split by rows with NO overlap.
    Wrapper function for chunk_with_row_protection.
    """
    # Wrapper function to pass force_chunking
    def split_with_protected_regions_wrapper(text, regions, chunk_size, chunk_overlap):
        return _split_with_protected_regions(text, regions, chunk_size, chunk_overlap, force_chunking)

    # Use unified table chunker that handles both HTML and Markdown
    return chunk_with_row_protection(
        text, chunk_size, chunk_overlap,
        split_with_protected_regions_wrapper, _chunk_table_unified
    )


# ============================================================================
# Internal Functions - Page/Line Mapping
# ============================================================================

def _extract_page_mapping(
    text: str,
    file_extension: str,
    page_tag_processor: Optional[Any] = None
) -> List[Dict[str, Any]]:
    """
    Extract page/slide mapping information from text.

    Recognizes page markers for various file formats:
    - PDF/PPT/DOCX: Page/slide markers
    - Excel: Sheet markers
    - Others: Line-based estimation

    Args:
        text: Original text
        file_extension: File extension
        page_tag_processor: PageTagProcessor instance for custom patterns

    Returns:
        Page mapping list [{"page_num": int, "start_pos": int, "end_pos": int, ...}, ...]
    """
    try:
        page_mapping: List[Dict[str, Any]] = []
        ext_lower = file_extension.lower() if file_extension else ""

        if ext_lower in ['pdf', 'ppt', 'pptx', 'doc', 'docx']:
            # Build patterns from PageTagProcessor or use defaults
            patterns = _get_page_marker_patterns(page_tag_processor)
            # Add OCR variants
            ocr_patterns = []
            for p in patterns:
                # Add (OCR) and (OCR+Ref) variants
                base_pattern = p.rstrip(']').rstrip(')')
                if base_pattern.endswith('\\d+'):
                    ocr_patterns.append(p[:-1] + r'\s*\(OCR\)\]')
                    ocr_patterns.append(p[:-1] + r'\s*\(OCR\+Ref\)\]')
            patterns.extend(ocr_patterns)

            for pattern in patterns:
                matches = list(re.finditer(pattern, text))
                if matches:
                    for i, match in enumerate(matches):
                        page_num = int(match.group(1))
                        # The marker line belongs to the page it opens
                        start = match.start()
                        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
                        page_mapping.append({
                            "page_num": page_num,
                            "start_pos": start,
                            "end_pos": end
                        })
                    page_mapping.sort(key=lambda x: x["page_num"])
                    break

            # Estimate pages for doc/docx if no markers found
            if not page_mapping and ext_lower in ['doc', 'docx']:
                chars_per_page = 1500
                text_len = len(text)
                if text_len > chars_per_page:
                    estimated_pages = (text_len + chars_per_page - 1) // chars_per_page
                    for page_num in range(1, estimated_pages + 1):
                        start = (page_num - 1) * chars_per_page
                        end = min(page_num * chars_per_page, text_len)
                        page_mapping.append({
                            "page_num": page_num,
                            "start_pos": start,
                            "end_pos": end
                        })

            if not page_mapping:
                page_mapping = [{"page_num": 1, "start_pos": 0, "end_pos": len(text)}]

        elif ext_lower in ['xlsx', 'xlsm', 'xltx', 'xltm', 'xls']:
            # Build sheet pattern from PageTagProcessor or use default
            sheet_pattern = _get_sheet_marker_pattern(page_tag_processor)
            matches = list(re.finditer(sheet_pattern, text))

            if matches:
                # Table chunks repeat the sheet marker, so the page number is the
                # sheet's order of first appearance rather than the match index.
                sheet_numbers: Dict[str, int] = {}
                for i, match in enumerate(matches):
                    sheet_name = match.group(1).strip()
                    page_num = sheet_numbers.setdefault(sheet_name, len(sheet_numbers) + 1)
                    start = match.start()
                    end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
                    page_mapping.append({
                        "page_num": page_num,
                        "start_pos": start,
                        "end_pos": end,
                        "sheet_name": sheet_name
                    })
            else:
                page_mapping = [{"page_num": 1, "start_pos": 0, "end_pos": len(text)}]

        else:
            # Line-based estimation for other file types
            lines = text.split('\n')
            lines_per_page = 1000

            if len(lines) > lines_per_page:
                page_count = (len(lines) + lines_per_page - 1) // lines_per_page
                current_pos = 0

                for page_num in range(1, page_count + 1):
                    start_line = (page_num - 1) * lines_per_page
                    end_line = min(page_num * lines_per_page, len(lines))
                    page_text = '\n'.join(lines[start_line:end_line])
                    start = current_pos
                    end = current_pos + len(page_text)
                    page_mapping.append({
                        "page_num": page_num,
                        "start_pos": start,
                        "end_pos": end
                    })
                    current_pos = end + 1
            else:
                page_mapping = [{"page_num": 1, "start_pos": 0, "end_pos": len(text)}]

        return page_mapping

    except Exception:
        return [{"page_num": 1, "start_pos": 0, "end_pos": len(text)}]


def _find_line_index_by_pos(
    pos: int,
    line_table: List[Dict[str, int]],
    starts: Optional[List[int]] = None
) -> int:
    """
    Find the line index corresponding to the given position.

    Args:
        pos: Position in text
        line_table: Line offset table
        starts: Precomputed line start offsets (pass it when calling in a loop)

    Returns:
        Line index (0-based)
    """
    try:
        if not line_table:
            return 0
        if starts is None:
            starts = [line["start"] for line in line_table]
        idx = bisect.bisect_right(starts, pos) - 1
        return 0 if idx < 0 else min(idx, len(line_table) - 1)
    except Exception:
        return 0


def _build_line_offset_table(
    text: str,
    file_extension: str,
    page_tag_processor: Optional[Any] = None
) -> List[Dict[str, int]]:
    """
    Build an offset table for each line in the text.

    Args:
        text: Original text
        file_extension: File extension
        page_tag_processor: PageTagProcessor instance for custom patterns

    Returns:
        Line offset table [{"line_num": int, "start": int, "end": int, "page": int}, ...]
    """
    try:
        lines = text.split('\n')
        table: List[Dict[str, int]] = []
        pos = 0
        page_mapping = _extract_page_mapping(text, file_extension, page_tag_processor)

        # Ranges sorted by start for a binary search (a linear scan per line was
        # quadratic: table chunks repeat the sheet marker, one range per chunk)
        ranges = sorted(page_mapping, key=lambda info: info["start_pos"])
        range_starts = [info["start_pos"] for info in ranges]

        def _page_for_pos(p: int) -> int:
            i = bisect.bisect_right(range_starts, p) - 1
            while i >= 0:
                info = ranges[i]
                if p < info["end_pos"]:
                    return info["page_num"]
                # Ranges do not overlap in practice; step back only past empty ones
                if info["start_pos"] < info["end_pos"]:
                    break
                i -= 1
            return 1

        for i, line in enumerate(lines):
            start = pos
            end = pos + len(line)
            mid = start + max(0, (end - start) // 2)
            page = _page_for_pos(mid)
            table.append({
                "line_num": i + 1,
                "start": start,
                "end": end,
                "page": page
            })
            pos = end + 1

        return table

    except Exception:
        return [{"line_num": 1, "start": 0, "end": len(text), "page": 1}]
