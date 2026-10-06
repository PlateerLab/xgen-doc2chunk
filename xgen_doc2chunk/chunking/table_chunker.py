# chunking_helper/table_chunker.py
"""
Table Chunker - Core table chunking logic

Main Features:
- Split large HTML tables to fit chunk_size
- Split large Markdown tables to fit chunk_size  
- Preserve and restore table structure (headers)
- rowspan/colspan aware splitting for HTML
- rowspan adjustment
- NO OVERLAP for table chunks (intentional to prevent data duplication)
"""
import logging
import re
from typing import Dict, List, Optional, Tuple

from xgen_doc2chunk.chunking.constants import (
    ParsedTable, TableRow, ParsedMarkdownTable, SpanningCell,
    TABLE_WRAPPER_OVERHEAD, CHUNK_INDEX_OVERHEAD,
    MARKDOWN_TABLE_SEPARATOR_PATTERN
)
from xgen_doc2chunk.chunking.table_parser import (
    parse_html_table, extract_cell_spans_with_positions, has_complex_spans,
    extract_spanning_cells
)

logger = logging.getLogger("document-processor")


def calculate_available_space(
    chunk_size: int,
    header_size: int,
    chunk_index: int = 0,
    total_chunks: int = 1
) -> int:
    """
    Calculate available space for data rows in a chunk.

    Args:
        chunk_size: Total chunk size
        header_size: Header size
        chunk_index: Current chunk index (0-based)
        total_chunks: Expected total number of chunks

    Returns:
        Number of characters available for data rows
    """
    # Fixed overhead
    overhead = TABLE_WRAPPER_OVERHEAD

    # Chunk index metadata overhead (only when total chunks > 1)
    if total_chunks > 1:
        overhead += CHUNK_INDEX_OVERHEAD

    # Header overhead (include header even for non-first chunks)
    overhead += header_size

    available = chunk_size - overhead

    return max(available, 100)  # Guarantee at least 100 characters


def adjust_rowspan_in_chunk(rows_html: List[str], total_rows_in_chunk: int) -> List[str]:
    """
    Readjust rowspan values for rows in a chunk.

    Adjusts rowspan values to match the number of rows included in the chunk
    so that the table renders correctly.

    Args:
        rows_html: List of HTML row strings included in the chunk
        total_rows_in_chunk: Total number of rows in the chunk

    Returns:
        List of HTML row strings with adjusted rowspan values
    """
    if not rows_html:
        return rows_html

    adjusted_rows = []

    for row_idx, row_html in enumerate(rows_html):
        remaining_rows = total_rows_in_chunk - row_idx

        def adjust_cell_rowspan(match):
            """Callback function to adjust cell rowspan"""
            tag = match.group(1)  # td or th
            attrs = match.group(2)
            content = match.group(3)

            # Extract current rowspan
            rowspan_match = re.search(r'rowspan=["\']?(\d+)["\']?', attrs, re.IGNORECASE)
            if rowspan_match:
                original_rowspan = int(rowspan_match.group(1))

                # Adjust if greater than remaining rows
                adjusted_rowspan = min(original_rowspan, remaining_rows)

                if adjusted_rowspan <= 1:
                    # Remove attribute if rowspan=1
                    new_attrs = re.sub(r'\s*rowspan=["\']?\d+["\']?', '', attrs, flags=re.IGNORECASE)
                else:
                    # Adjust rowspan value
                    new_attrs = re.sub(
                        r'rowspan=["\']?\d+["\']?',
                        f"rowspan='{adjusted_rowspan}'",
                        attrs,
                        flags=re.IGNORECASE
                    )

                return f'<{tag}{new_attrs}>{content}</{tag}>'

            return match.group(0)

        # Cell pattern: <td ...>...</td> or <th ...>...</th>
        cell_pattern = r'<(td|th)([^>]*)>(.*?)</\1>'
        adjusted_row = re.sub(cell_pattern, adjust_cell_rowspan, row_html, flags=re.DOTALL | re.IGNORECASE)

        adjusted_rows.append(adjusted_row)

    return adjusted_rows


def build_table_chunk(
    header_html: str,
    data_rows: List[TableRow],
    chunk_index: int = 0,
    total_chunks: int = 1,
    context_prefix: str = ""
) -> str:
    """
    Build a complete table HTML for a chunk.

    Automatically adjusts rowspan if it exceeds the chunk boundary.

    Args:
        header_html: HTML of header rows
        data_rows: Data rows
        chunk_index: Current chunk index (0-based)
        total_chunks: Total number of chunks
        context_prefix: Context info (metadata, sheet info, etc.) - included in all chunks

    Returns:
        Complete table HTML
    """
    parts = []

    # Context info (metadata, sheet info, etc.) - included in all chunks
    if context_prefix:
        parts.append(context_prefix)

    # Chunk index metadata (only when more than 1 chunk)
    if total_chunks > 1:
        parts.append(f"[Table Chunk {chunk_index + 1}/{total_chunks}]")

    # Table start
    parts.append("<table border='1'>")

    # Header (if exists)
    if header_html:
        parts.append(header_html)

    # Extract HTML for data rows
    rows_html = [row.html for row in data_rows]

    # Adjust rowspan
    adjusted_rows = adjust_rowspan_in_chunk(rows_html, len(data_rows))

    # Add adjusted rows
    for row_html in adjusted_rows:
        parts.append(row_html)

    # Table end
    parts.append("</table>")

    return "\n".join(parts)


def update_chunk_metadata(chunks: List[str], total_chunks: int) -> List[str]:
    """
    Update chunk metadata (total chunk count).
    """
    updated_chunks = []

    for idx, chunk in enumerate(chunks):
        # Existing metadata pattern
        old_pattern = r'\[Table Chunk \d+/\d+\]'
        new_metadata = f"[Table Chunk {idx + 1}/{total_chunks}]"

        if re.search(old_pattern, chunk):
            updated_chunk = re.sub(old_pattern, new_metadata, chunk)
        else:
            # Add metadata if not present
            updated_chunk = f"{new_metadata}\n{chunk}"

        updated_chunks.append(updated_chunk)

    return updated_chunks


def split_table_into_chunks(
    parsed_table: ParsedTable,
    chunk_size: int,
    chunk_overlap: int = 0,
    context_prefix: str = ""
) -> List[str]:
    """
    Split a parsed table to fit chunk_size.
    Each chunk has a complete table structure (including headers).

    NOTE: Table chunking does NOT apply overlap.
    Data duplication degrades search quality, so overlap is intentionally excluded.

    Row splitting rules:
    - Minimum 1 row per chunk (rows are NEVER split)
    - Rows are added while they fit in chunk_size
    - Only exceeds chunk_size when a single row is larger than the budget

    Args:
        parsed_table: Parsed table information
        chunk_size: Maximum chunk size
        chunk_overlap: Not used (kept for compatibility)
        context_prefix: Context info (metadata, sheet info, etc.) - included in all chunks

    Returns:
        List of split table HTML chunks
    """
    data_rows = parsed_table.data_rows
    header_html = parsed_table.header_html
    header_size = parsed_table.header_size

    # Calculate context size
    context_size = len(context_prefix) + 2 if context_prefix else 0  # Including newline

    if not data_rows:
        # Return original if no data rows
        return [parsed_table.original_html]

    # Calculate estimated chunk count (approximate)
    total_data_size = sum(row.char_length for row in data_rows)
    available_per_chunk = calculate_available_space(chunk_size, header_size + context_size, 0, 1)
    estimated_chunks = max(1, (total_data_size + available_per_chunk - 1) // available_per_chunk)

    # Recalculate with actual chunk count
    available_per_chunk = calculate_available_space(chunk_size, header_size + context_size, 0, estimated_chunks)

    chunks: List[str] = []
    current_rows: List[TableRow] = []
    current_size = 0
    # Table chunking does not apply overlap (prevent data duplication)

    for row_idx, row in enumerate(data_rows):
        row_size = row.char_length + 1  # Including newline

        # Flush when this row no longer fits. Chunks used to be filled up to
        # 1.5x chunk_size here, which made almost every table chunk oversized.
        if current_rows and (current_size + row_size > available_per_chunk):
            chunk_html = build_table_chunk(
                header_html,
                current_rows,
                chunk_index=len(chunks),
                total_chunks=estimated_chunks,
                context_prefix=context_prefix
            )
            chunks.append(chunk_html)

            # Start new chunk with this row (minimum 1 row guaranteed)
            current_rows = [row]
            current_size = row_size
        else:
            # Row fits - add to current chunk
            current_rows.append(row)
            current_size += row_size

    # Process last chunk
    if current_rows:
        chunk_html = build_table_chunk(
            header_html,
            current_rows,
            chunk_index=len(chunks),
            total_chunks=max(len(chunks) + 1, estimated_chunks),
            context_prefix=context_prefix
        )
        chunks.append(chunk_html)

    # Update metadata with actual total chunk count
    if len(chunks) != estimated_chunks and len(chunks) > 1:
        chunks = update_chunk_metadata(chunks, len(chunks))

    logger.info(f"Table split into {len(chunks)} chunks (original: {len(parsed_table.original_html)} chars)")

    return chunks


def _compute_carried_cells(
    data_rows: List[TableRow]
) -> List[Dict[int, Tuple[SpanningCell, int]]]:
    """
    For each row, snapshot the spanning cells that cover the row but physically
    live in an earlier row (i.e., cells that would be lost if a chunk started here).

    Uses the same decrement/update rules as the block identification loop in
    split_table_preserving_rowspan so both stay consistent.

    Args:
        data_rows: Table data rows

    Returns:
        Per-row dict {column_position: (SpanningCell, remaining_rows)} where
        remaining_rows includes the current row.
    """
    carried_per_row: List[Dict[int, Tuple[SpanningCell, int]]] = []
    active: Dict[int, Tuple[SpanningCell, int]] = {}  # col -> (cell, remaining including current row)

    for row_idx, row in enumerate(data_rows):
        # Decrease remaining rowspan from previous row (except first row)
        if row_idx > 0:
            finished_cols = []
            for col in list(active.keys()):
                cell, remaining = active[col]
                remaining -= 1
                if remaining <= 0:
                    finished_cols.append(col)
                else:
                    active[col] = (cell, remaining)
            for col in finished_cols:
                del active[col]

        # Snapshot BEFORE adding this row's own spans:
        # exactly the cells that started earlier and still cover this row
        carried_per_row.append(dict(active))

        # Add new rowspans starting from current row (longer span takes priority)
        new_cells = extract_spanning_cells(row.html)
        for col, cell in new_cells.items():
            if col not in active or cell.rowspan > active[col][1]:
                active[col] = (cell, cell.rowspan)

    return carried_per_row


def reissue_carried_cells(
    row_html: str,
    carried: Dict[int, Tuple[SpanningCell, int]]
) -> str:
    """
    Re-issue carried spanning cells into a row that starts a continuation chunk.

    Inserts each carried cell at its column position with rowspan set to its
    remaining coverage (build_table_chunk clamps it to the chunk's row count).

    Args:
        row_html: Row HTML (first row of a continuation chunk)
        carried: {column_position: (SpanningCell, remaining_rows)}

    Returns:
        Row HTML with carried cells inserted
    """
    if not carried:
        return row_html

    tr_match = re.match(r'\s*<tr[^>]*>(.*)</tr>\s*$', row_html, re.DOTALL | re.IGNORECASE)
    inner = tr_match.group(1) if tr_match else row_html

    cell_pattern = r'<(th|td)([^>]*)>(.*?)</\1>'
    existing_cells = list(re.finditer(cell_pattern, inner, re.DOTALL | re.IGNORECASE))

    def render(cell: SpanningCell, remaining: int) -> str:
        attrs = f" rowspan='{remaining}'"
        if cell.attrs:
            attrs += f" {cell.attrs}"
        return f"<{cell.tag}{attrs}>{cell.content}</{cell.tag}>"

    pending = sorted(carried.items())  # by column position
    parts: List[str] = []
    pending_idx = 0
    current_col = 0

    for match in existing_cells:
        # Insert carried cells occupying column slots before this existing cell
        while pending_idx < len(pending) and pending[pending_idx][0] <= current_col:
            _, (cell, remaining) = pending[pending_idx]
            parts.append(render(cell, remaining))
            current_col += cell.colspan
            pending_idx += 1

        parts.append(match.group(0))
        colspan_match = re.search(r'colspan=["\']?(\d+)["\']?', match.group(2), re.IGNORECASE)
        current_col += int(colspan_match.group(1)) if colspan_match else 1

    # Carried cells positioned after all existing cells
    while pending_idx < len(pending):
        _, (cell, remaining) = pending[pending_idx]
        parts.append(render(cell, remaining))
        pending_idx += 1

    return f"<tr>{''.join(parts)}</tr>"


def split_oversized_block(
    block_rows: List[TableRow],
    block_carried: List[Dict[int, Tuple[SpanningCell, int]]],
    header_html: str,
    available_space: int,
    max_chunk_data_size: int,
    context_prefix: str,
    start_chunk_index: int
) -> List[str]:
    """
    Split a single rowspan block that exceeds max_chunk_data_size into
    row-level sub-chunks of at most available_space. Rows are NEVER split internally.

    Each sub-chunk that does not start at the block's first row gets the
    active spanning cells re-issued into its first row, so every sub-chunk
    remains a self-contained, structurally valid table.

    Args:
        block_rows: Rows of the oversized block
        block_carried: Carried-cell snapshot for each row (aligned with block_rows)
        header_html: Header HTML restored in every chunk
        available_space: Target data size per chunk
        max_chunk_data_size: Size above which a block counts as oversized (kept for compatibility)
        context_prefix: Context info included in all chunks
        start_chunk_index: Chunk index offset for metadata
    Returns:
        List of table chunk HTML strings
    """
    chunks: List[str] = []
    current_rows: List[TableRow] = []
    current_size = 0

    for i, row in enumerate(block_rows):
        row_size = row.char_length + 1  # Including newline

        # Flush when the row no longer fits in the chunk budget
        # (rows themselves are never split - a single oversized row
        #  becomes its own chunk)
        if current_rows and current_size + row_size > available_space:
            chunks.append(build_table_chunk(
                header_html, current_rows,
                start_chunk_index + len(chunks),
                start_chunk_index + len(chunks) + 2,
                context_prefix=context_prefix
            ))
            current_rows = []
            current_size = 0

        # A sub-chunk starting mid-block loses the spanning cells that live in
        # earlier rows - re-issue them into its first row
        if not current_rows and i > 0 and block_carried[i]:
            new_html = reissue_carried_cells(row.html, block_carried[i])
            row = TableRow(
                html=new_html,
                is_header=row.is_header,
                cell_count=row.cell_count + len(block_carried[i]),
                char_length=len(new_html)
            )

        current_rows.append(row)
        current_size += row.char_length + 1

    if current_rows:
        chunks.append(build_table_chunk(
            header_html, current_rows,
            start_chunk_index + len(chunks),
            start_chunk_index + len(chunks) + 2,
            context_prefix=context_prefix
        ))

    return chunks


def split_table_preserving_rowspan(
    parsed_table: ParsedTable,
    chunk_size: int,
    chunk_overlap: int,
    context_prefix: str = "",
    force_chunking: bool = False
) -> List[str]:
    """
    Split a table considering rowspan.

    Rows connected by rowspan are kept together as semantic blocks.

    NOTE: Table chunking does NOT apply overlap.
    Data duplication degrades search quality, so overlap is intentionally excluded.

    Algorithm:
    1. Track active rowspan for each row (by column position, considering colspan)
    2. If all rowspans from previous row end and new rowspan starts, create new block
    3. Combine blocks to fit chunk_size (a block larger than chunk_size gets its own chunk)
    4. (force_chunking only) A block exceeding 1.5x chunk_size is split internally
       at row boundaries, with carried spanning cells re-issued in each sub-chunk

    Args:
        parsed_table: Parsed table
        chunk_size: Chunk size
        chunk_overlap: Not used (kept for compatibility)
        context_prefix: Context info (metadata, sheet info, etc.)
        force_chunking: Allow splitting inside oversized rowspan blocks.
            False (default) keeps the previous behavior: blocks are never
            split internally, regardless of size.

    Returns:
        List of split table chunks
    """
    data_rows = parsed_table.data_rows
    header_html = parsed_table.header_html
    header_size = parsed_table.header_size

    # Calculate context size
    context_size = len(context_prefix) + 2 if context_prefix else 0

    if not data_rows:
        if context_prefix:
            return [f"{context_prefix}\n{parsed_table.original_html}"]
        return [parsed_table.original_html]

    # === Identify rowspan blocks ===
    # Block = group of consecutive rows connected by rowspan
    active_rowspans: Dict[int, int] = {}  # column_position -> remaining_rows (including current row)
    row_block_ids: List[int] = []  # Block ID for each row
    current_block_id = -1

    for row_idx, row in enumerate(data_rows):
        # 1. Decrease remaining rowspan from previous row (except first row)
        if row_idx > 0:
            finished_cols = []
            for col in list(active_rowspans.keys()):
                active_rowspans[col] -= 1
                if active_rowspans[col] <= 0:
                    finished_cols.append(col)
            for col in finished_cols:
                del active_rowspans[col]

        # State after decrease (before adding new spans)
        had_active_before_new = len(active_rowspans) > 0

        # 2. Add new rowspans starting from current row
        new_spans = extract_cell_spans_with_positions(row.html)
        for col, span in new_spans.items():
            # Update if larger than existing rowspan (longer span takes priority)
            if col not in active_rowspans or span > active_rowspans[col]:
                active_rowspans[col] = span

        has_active_now = len(active_rowspans) > 0
        has_new_span = len(new_spans) > 0

        # Block determination logic:
        # - No active rowspan -> independent block
        # - No active after previous row processing but new span starts -> new block
        # - Otherwise maintain existing block
        if not has_active_now:
            # No rowspan - independent row
            current_block_id += 1
            row_block_ids.append(current_block_id)
        elif not had_active_before_new and has_new_span:
            # All previous rowspans ended and new rowspan starts - new block
            current_block_id += 1
            row_block_ids.append(current_block_id)
        else:
            # Maintain existing block
            row_block_ids.append(current_block_id)

    # Group rows by block
    block_groups: Dict[int, List[int]] = {}
    for row_idx, block_id in enumerate(row_block_ids):
        if block_id not in block_groups:
            block_groups[block_id] = []
        block_groups[block_id].append(row_idx)

    # Create row_groups in sorted block order
    row_groups: List[List[int]] = [
        block_groups[block_id]
        for block_id in sorted(block_groups.keys())
    ]

    # === Combine groups into chunks ===
    chunks: List[str] = []
    current_rows: List[TableRow] = []
    current_size = 0

    available_space = calculate_available_space(chunk_size, header_size + context_size, 0, 1)
    # Blocks above 1.5x chunk_size count as oversized (split only with force_chunking)
    max_chunk_data_size = int(chunk_size * 1.5) - header_size - context_size - CHUNK_INDEX_OVERHEAD

    # Carried-cell snapshots are only needed when oversized blocks may be split
    carried_per_row = _compute_carried_cells(data_rows) if force_chunking else None

    for group in row_groups:
        group_rows = [data_rows[idx] for idx in group]
        group_size = sum(row.char_length + 1 for row in group_rows)

        if force_chunking and group_size > max_chunk_data_size:
            # Oversized rowspan block: cannot fit in a single chunk.
            # Flush the current chunk, then split the block internally at row
            # boundaries with carried spanning cells re-issued (rows stay intact).
            if current_rows:
                chunks.append(build_table_chunk(
                    header_html, current_rows, len(chunks), len(chunks) + 2,
                    context_prefix=context_prefix
                ))
                current_rows = []
                current_size = 0

            block_carried = [carried_per_row[idx] for idx in group]
            chunks.extend(split_oversized_block(
                group_rows, block_carried, header_html,
                available_space, max_chunk_data_size, context_prefix, len(chunks)
            ))
            continue

        if current_rows and current_size + group_size > available_space:
            # Flush when the group no longer fits (no 1.5x fill)
            chunks.append(build_table_chunk(
                header_html, current_rows, len(chunks), len(chunks) + 2,
                context_prefix=context_prefix
            ))
            current_rows = group_rows[:]
            current_size = group_size
        else:
            current_rows.extend(group_rows)
            current_size += group_size

    # Last chunk
    if current_rows:
        chunks.append(build_table_chunk(
            header_html, current_rows, len(chunks), len(chunks) + 1,
            context_prefix=context_prefix
        ))

    # Update chunk count
    if len(chunks) > 1:
        chunks = update_chunk_metadata(chunks, len(chunks))

    return chunks


def chunk_large_table(
    table_html: str,
    chunk_size: int,
    chunk_overlap: int,
    context_prefix: str = "",
    force_chunking: bool = False
) -> List[str]:
    """
    Split large HTML table to fit chunk_size.
    Restores table structure (headers) in each chunk.

    Also handles complex tables with rowspan.

    NOTE: Table chunking does NOT apply overlap.
    Data duplication degrades search quality, so overlap is intentionally excluded.

    Args:
        table_html: HTML table string
        chunk_size: Maximum chunk size
        chunk_overlap: Not used (kept for compatibility)
        context_prefix: Context info (metadata, sheet info, etc.) - included in all chunks
        force_chunking: Allow splitting inside oversized rowspan blocks
            (blocks whose rows are tied together by merged cells). Carried
            spanning cells are re-issued so each chunk stays self-contained.
            False (default) keeps the previous behavior.

    Returns:
        List of split table HTML chunks
    """
    # Parse table
    parsed = parse_html_table(table_html)

    if not parsed:
        logger.warning("Failed to parse table, returning original")
        if context_prefix:
            return [f"{context_prefix}\n{table_html}"]
        return [table_html]

    # No need to split if table fits in chunk_size
    if len(table_html) + len(context_prefix) <= chunk_size:
        if context_prefix:
            return [f"{context_prefix}\n{table_html}"]
        return [table_html]

    # No need to split if no data rows
    if not parsed.data_rows:
        if context_prefix:
            return [f"{context_prefix}\n{table_html}"]
        return [table_html]

    # Check for complex spans (rowspan)
    if has_complex_spans(table_html):
        logger.info("Complex table with rowspan detected, using span-aware splitting")
        return split_table_preserving_rowspan(
            parsed, chunk_size, chunk_overlap, context_prefix,
            force_chunking=force_chunking
        )

    # Standard table splitting
    chunks = split_table_into_chunks(parsed, chunk_size, chunk_overlap, context_prefix)

    return chunks


# ============================================================================
# Markdown Table Chunking Functions
# ============================================================================

def parse_markdown_table(table_text: str) -> Optional[ParsedMarkdownTable]:
    """
    Parse a Markdown table and extract structural information.

    A Markdown table has:
    - Header row: | col1 | col2 | col3 |
    - Separator row: |---|---|---| or |:---:|:---|---:|
    - Data rows: | data1 | data2 | data3 |

    Args:
        table_text: Markdown table text

    Returns:
        ParsedMarkdownTable object or None if parsing fails
    """
    try:
        # Split into lines and filter empty lines
        lines = [line.strip() for line in table_text.strip().split('\n') if line.strip()]

        if len(lines) < 2:
            logger.debug("Not enough lines for a valid Markdown table")
            return None

        # Find header and separator rows
        header_row = None
        separator_row = None
        separator_idx = -1

        for idx, line in enumerate(lines):
            # Check if this line is a separator (contains only |, -, :, and spaces)
            if re.match(MARKDOWN_TABLE_SEPARATOR_PATTERN, line):
                separator_row = line
                separator_idx = idx
                # Header is the line before separator
                if idx > 0:
                    header_row = lines[idx - 1]
                break

        if not separator_row or not header_row:
            # Try simpler detection: first row is header, second row is separator
            if len(lines) >= 2 and lines[0].startswith('|') and '---' in lines[1]:
                header_row = lines[0]
                separator_row = lines[1]
                separator_idx = 1
            else:
                logger.debug("Could not identify header/separator in Markdown table")
                return None

        # Count columns from separator
        total_cols = separator_row.count('|') - 1  # -1 because |---|---| has n+1 pipes for n columns

        # Data rows are all rows after separator
        data_rows = lines[separator_idx + 1:]

        # Lines before the header row (e.g. "[Table 2] 2026 sales") used to be
        # dropped when the table was split. Keep them and repeat them in every chunk.
        preamble = "\n".join(lines[:max(separator_idx - 1, 0)])

        # Construct header text (header + separator) for restoration in each chunk
        header_text = f"{header_row}\n{separator_row}"
        header_size = len(header_text) + 1  # +1 for newline
        if preamble:
            header_size += len(preamble) + 1

        return ParsedMarkdownTable(
            header_row=header_row,
            separator_row=separator_row,
            data_rows=data_rows,
            total_cols=total_cols,
            original_text=table_text,
            header_text=header_text,
            header_size=header_size,
            preamble=preamble
        )

    except Exception as e:
        logger.warning(f"Failed to parse Markdown table: {e}")
        return None


def build_markdown_table_chunk(
    header_text: str,
    data_rows: List[str],
    chunk_index: int = 0,
    total_chunks: int = 1,
    context_prefix: str = "",
    preamble: str = ""
) -> str:
    """
    Build a complete Markdown table chunk with header restored.

    Args:
        header_text: Header row + separator row
        data_rows: List of data row strings
        chunk_index: Current chunk index (0-based)
        total_chunks: Total number of chunks
        context_prefix: Context info (metadata, sheet info, etc.) - included in all chunks
        preamble: Lines that preceded the header (table marker, caption) - included in all chunks

    Returns:
        Complete Markdown table chunk
    """
    parts = []

    # Add context prefix if provided
    if context_prefix:
        parts.append(context_prefix)

    if preamble:
        parts.append(preamble)

    # Add chunk index metadata (only if more than 1 chunk)
    if total_chunks > 1:
        parts.append(f"[Table Chunk {chunk_index + 1}/{total_chunks}]")

    # Add header (header row + separator row)
    parts.append(header_text)

    # Add data rows
    for row in data_rows:
        parts.append(row)

    return "\n".join(parts)


def update_markdown_chunk_metadata(chunks: List[str], total_chunks: int) -> List[str]:
    """
    Update chunk metadata (total chunk count) in Markdown table chunks.

    Args:
        chunks: List of chunks
        total_chunks: Actual total number of chunks

    Returns:
        Updated chunks with correct metadata
    """
    updated_chunks = []

    for idx, chunk in enumerate(chunks):
        # Pattern for existing metadata
        old_pattern = r'\[Table Chunk \d+/\d+\]'
        new_metadata = f"[Table Chunk {idx + 1}/{total_chunks}]"

        if re.search(old_pattern, chunk):
            updated_chunk = re.sub(old_pattern, new_metadata, chunk)
        else:
            # No metadata found - add it
            updated_chunk = f"{new_metadata}\n{chunk}"

        updated_chunks.append(updated_chunk)

    return updated_chunks


# A row longer than this many chunk_size is split by its longest cell
LONG_ROW_FACTOR = 3

_UNESCAPED_PIPE = re.compile(r'(?<!\\)\|')


def _cut_text(text: str, size: int) -> List[str]:
    """Cut text into pieces of at most size characters, preferring whitespace."""
    pieces: List[str] = []
    while len(text) > size:
        cut = text.rfind(" ", size // 2, size + 1)
        if cut <= 0:
            cut = size
        pieces.append(text[:cut].strip())
        text = text[cut:].strip()
    if text:
        pieces.append(text)
    return pieces


def split_long_markdown_row(row: str, budget: int) -> List[str]:
    """
    Split one oversized Markdown row into several full-width rows.

    The longest cell is cut into pieces and every piece gets its own copy of the
    row, so each piece keeps the other columns (id, title, ...) and every row
    still has the full column count. A row whose length is not dominated by a
    single cell (a very wide table) is returned unchanged.

    Args:
        row: Markdown table row (| a | b | c |)
        budget: Target row length

    Returns:
        List of rows
    """
    stripped = row.strip()
    if not (stripped.startswith('|') and stripped.endswith('|')):
        return [row]
    cells = [c.strip() for c in _UNESCAPED_PIPE.split(stripped)[1:-1]]
    if not cells:
        return [row]
    longest = max(range(len(cells)), key=lambda i: len(cells[i]))
    rest = len(stripped) - len(cells[longest])
    if len(cells[longest]) < len(stripped) // 2:
        return [row]
    piece_size = max(budget - rest, 200)
    pieces = _cut_text(cells[longest], piece_size)
    # Never leave an escaped pipe split from its backslash
    pieces = [p[:-1] if p.endswith('\\') else p for p in pieces]
    if len(pieces) <= 1:
        return [row]
    rows = []
    for piece in pieces:
        parts = list(cells)
        parts[longest] = piece
        rows.append("| " + " | ".join(parts) + " |")
    return rows


def split_markdown_table_into_chunks(
    parsed_table: ParsedMarkdownTable,
    chunk_size: int,
    chunk_overlap: int = 0,
    context_prefix: str = ""
) -> List[str]:
    """
    Split a parsed Markdown table into chunks that fit chunk_size.
    Each chunk is a complete Markdown table with headers (and the preamble,
    such as a table caption) restored.

    Rows are added while they fit in chunk_size. A single row larger than the
    budget gets its own chunk; a row larger than LONG_ROW_FACTOR x chunk_size is
    split by its longest cell (see split_long_markdown_row).

    NOTE: Table chunking does NOT apply overlap.
    Data duplication degrades search quality, so overlap is intentionally excluded.

    Args:
        parsed_table: Parsed Markdown table information
        chunk_size: Maximum chunk size
        chunk_overlap: Not used (kept for compatibility)
        context_prefix: Context info (metadata, sheet info, etc.) - included in all chunks

    Returns:
        List of Markdown table chunk strings
    """
    data_rows = parsed_table.data_rows
    header_text = parsed_table.header_text
    header_size = parsed_table.header_size
    preamble = parsed_table.preamble

    # Calculate context size
    context_size = len(context_prefix) + 2 if context_prefix else 0  # +2 for newline

    if not data_rows:
        # No data rows - return original
        if context_prefix:
            return [f"{context_prefix}\n{parsed_table.original_text}"]
        return [parsed_table.original_text]

    # Calculate available space per chunk
    # Overhead: chunk index metadata (~25 chars) + header + context
    available_per_chunk = chunk_size - header_size - context_size - CHUNK_INDEX_OVERHEAD

    long_row_limit = chunk_size * LONG_ROW_FACTOR
    if any(len(row) > long_row_limit for row in data_rows):
        row_budget = max(available_per_chunk, chunk_size // 2)
        expanded: List[str] = []
        for row in data_rows:
            if len(row) > long_row_limit:
                expanded.extend(split_long_markdown_row(row, row_budget))
            else:
                expanded.append(row)
        data_rows = expanded

    estimated_chunks = 1
    total_data_size = sum(len(row) + 1 for row in data_rows)  # +1 for newline
    if available_per_chunk > 0:
        estimated_chunks = max(1, (total_data_size + available_per_chunk - 1) // available_per_chunk)

    chunks: List[str] = []
    current_rows: List[str] = []
    current_size = 0

    for row in data_rows:
        row_size = len(row) + 1  # +1 for newline

        # Flush when this row no longer fits. Chunks used to be filled up to
        # 1.5x chunk_size here, which made almost every table chunk oversized.
        if current_rows and (current_size + row_size > available_per_chunk):
            chunk_text = build_markdown_table_chunk(
                header_text,
                current_rows,
                chunk_index=len(chunks),
                total_chunks=estimated_chunks,
                context_prefix=context_prefix,
                preamble=preamble
            )
            chunks.append(chunk_text)

            # Start new chunk with this row (minimum 1 row guaranteed)
            current_rows = [row]
            current_size = row_size
        else:
            # Row fits - add to current chunk
            current_rows.append(row)
            current_size += row_size

    # Handle last chunk
    if current_rows:
        chunk_text = build_markdown_table_chunk(
            header_text,
            current_rows,
            chunk_index=len(chunks),
            total_chunks=max(len(chunks) + 1, estimated_chunks),
            context_prefix=context_prefix,
            preamble=preamble
        )
        chunks.append(chunk_text)

    # Update total chunk count in metadata if different from estimate
    if len(chunks) != estimated_chunks and len(chunks) > 1:
        chunks = update_markdown_chunk_metadata(chunks, len(chunks))
    elif len(chunks) == 1 and estimated_chunks > 1:
        chunks = [re.sub(r'\[Table Chunk \d+/\d+\]\n', '', chunks[0], count=1)]

    logger.info(f"Markdown table split into {len(chunks)} chunks (original: {len(parsed_table.original_text)} chars)")

    return chunks


def chunk_large_markdown_table(
    table_text: str,
    chunk_size: int,
    chunk_overlap: int,
    context_prefix: str = ""
) -> List[str]:
    """
    Split a large Markdown table to fit chunk_size.
    Restores table structure (header + separator) in each chunk.

    NOTE: Table chunking does NOT apply overlap.
    Data duplication degrades search quality, so overlap is intentionally excluded.

    Args:
        table_text: Markdown table text
        chunk_size: Maximum chunk size
        chunk_overlap: Not used (kept for compatibility)
        context_prefix: Context info (metadata, sheet info, etc.) - included in all chunks

    Returns:
        List of split Markdown table chunks
    """
    # Parse table
    parsed = parse_markdown_table(table_text)

    if not parsed:
        logger.warning("Failed to parse Markdown table, returning original")
        if context_prefix:
            return [f"{context_prefix}\n{table_text}"]
        return [table_text]

    # No need to split if table fits in chunk_size
    if len(table_text) + len(context_prefix) <= chunk_size:
        if context_prefix:
            return [f"{context_prefix}\n{table_text}"]
        return [table_text]

    # No need to split if no data rows
    if not parsed.data_rows:
        if context_prefix:
            return [f"{context_prefix}\n{table_text}"]
        return [table_text]

    # Split table into chunks
    chunks = split_markdown_table_into_chunks(parsed, chunk_size, chunk_overlap, context_prefix)

    return chunks


def is_markdown_table(text: str) -> bool:
    """
    Check if text is a Markdown table.

    A Markdown table has:
    - Lines starting with |
    - A separator line with |---|

    Args:
        text: Text to check

    Returns:
        True if text is a Markdown table
    """
    lines = text.strip().split('\n')
    if len(lines) < 2:
        return False

    # Check for | at start of lines and separator pattern
    has_pipe_rows = any(line.strip().startswith('|') for line in lines)
    has_separator = any('---' in line and '|' in line for line in lines)

    return has_pipe_rows and has_separator


# Note: detect_table_type and chunk_large_table_unified were removed because they
# were not referenced anywhere in the codebase and duplicated logic handled elsewhere
# (e.g., via _chunk_table_unified in chunking.py). Keeping a single authoritative
# implementation reduces the risk of divergent behavior.
