# your_package/document_processor/excel_handler.py
"""
Excel Handler - Excel Document Processor (XLSX/XLS)

Main Features:
- Metadata extraction (title, author, subject, keywords, creation date, modification date, etc.)
- Whole used range of every sheet (no row/column cap), display formats applied
  (dates, percentages, thousands separators, currency), formulas without cached values kept as text
- Table detection that survives spacer columns/rows, form layouts (label: value) and sparse cells
- Markdown tables with flattened multi-row (merged) headers and table titles, so every chunk of a
  split table carries its header (see excel_helper/sheet_layout.py)
- Large workbooks are streamed (openpyxl read_only) to bound memory
- Inline image extraction and local storage, chart processing, textboxes
- Multi-sheet support; xlsx / xlsm / xltx / xltm / xls

Options (DocumentProcessor config["spreadsheet"]):
- include_hidden (default True): keep hidden sheets, rows and columns
- streaming_threshold_bytes (default 20 MiB of sheet XML): stream larger workbooks

Class-based Handler:
- ExcelHandler class inherits from BaseHandler to manage config/image_processor
"""
from __future__ import annotations

import logging
import os
from typing import TYPE_CHECKING, Any, Dict, List, Optional, Set

from xgen_doc2chunk.core.processor.base_handler import BaseHandler
from xgen_doc2chunk.core.functions.img_processor import ImageProcessor
from xgen_doc2chunk.core.functions.chart_extractor import BaseChartExtractor
from xgen_doc2chunk.core.processor.excel_helper.excel_chart_extractor import ExcelChartExtractor

if TYPE_CHECKING:
    from openpyxl.workbook import Workbook
    from openpyxl.worksheet.worksheet import Worksheet
    from xgen_doc2chunk.core.document_processor import CurrentFile
from xgen_doc2chunk.core.processor.excel_helper import (
    # Textbox
    extract_textboxes_from_xlsx,
    extract_textboxes_from_xls,
)
from xgen_doc2chunk.core.processor.excel_helper.excel_metadata import (
    XLSXMetadataExtractor,
    XLSMetadataExtractor,
)
from xgen_doc2chunk.core.processor.excel_helper.excel_image_processor_xlsx import (
    ExcelImageProcessor,
)
from xgen_doc2chunk.core.processor.excel_helper.excel_image_processor_xls import (
    XLSImageProcessor,
)
from xgen_doc2chunk.core.processor.excel_helper.sheet_grid import (
    STREAMING_THRESHOLD_BYTES,
    SheetGrid,
    grid_from_html_rows,
    grid_from_xlrd,
    grids_from_workbook,
    is_large_workbook,
    load_streaming_workbook,
)
from xgen_doc2chunk.core.processor.excel_helper.sheet_layout import render_sheet

#: openpyxl 이 읽는 xlsx 계열(매크로 통합 문서·서식 파일은 같은 OOXML 이다).
XLSX_FAMILY = ("xlsx", "xlsm", "xltx", "xltm")

logger = logging.getLogger("document-processor")


# ============================================================================
# ExcelHandler Class
# ============================================================================

class ExcelHandler(BaseHandler):
    """
    Excel Document Handler (XLSX/XLS)

    Inherits from BaseHandler to manage config and image_processor at instance level.

    Usage:
        handler = ExcelHandler(config=config, image_processor=image_processor)
        text = handler.extract_text(current_file)
    """

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._xlsx_metadata_extractor = None
        self._xls_metadata_extractor = None

    def _create_file_converter(self):
        """Create Excel-specific file converter."""
        from xgen_doc2chunk.core.processor.excel_helper.excel_file_converter import ExcelFileConverter
        return ExcelFileConverter()

    def _create_preprocessor(self):
        """Create Excel-specific preprocessor."""
        from xgen_doc2chunk.core.processor.excel_helper.excel_preprocessor import ExcelPreprocessor
        return ExcelPreprocessor()

    def _create_chart_extractor(self) -> BaseChartExtractor:
        """Create Excel-specific chart extractor."""
        return ExcelChartExtractor(self._chart_processor)

    def _create_metadata_extractor(self):
        """Create XLSX-specific metadata extractor (default)."""
        return XLSXMetadataExtractor()

    def _create_format_image_processor(self):
        """Create Excel-specific image processor (XLSX)."""
        return ExcelImageProcessor(
            directory_path=self._image_processor.config.directory_path,
            tag_prefix=self._image_processor.config.tag_prefix,
            tag_suffix=self._image_processor.config.tag_suffix,
            storage_backend=self._image_processor.storage_backend,
        )

    def _create_xls_image_processor(self):
        """Create XLS-specific image processor."""
        return XLSImageProcessor(
            directory_path=self._image_processor.config.directory_path,
            tag_prefix=self._image_processor.config.tag_prefix,
            tag_suffix=self._image_processor.config.tag_suffix,
            storage_backend=self._image_processor.storage_backend,
        )

    def _get_xls_metadata_extractor(self):
        """Get XLS-specific metadata extractor."""
        if self._xls_metadata_extractor is None:
            self._xls_metadata_extractor = XLSMetadataExtractor()
        return self._xls_metadata_extractor

    def extract_text(
        self,
        current_file: "CurrentFile",
        extract_metadata: bool = True,
        **kwargs
    ) -> str:
        """
        Extract text from Excel file.

        Args:
            current_file: CurrentFile dict containing file info and binary data
            extract_metadata: Whether to extract metadata
            **kwargs: Additional options

        Returns:
            Extracted text
        """
        file_path = current_file.get("file_path", "unknown")
        ext = current_file.get("file_extension", os.path.splitext(file_path)[1]).lower()
        # Normalize extension (remove leading dot if present)
        ext = ext.lstrip('.')
        self.logger.info(f"Excel processing: {file_path}, ext: {ext}")

        # Check if file content is actually HTML disguised as Excel
        file_data = current_file.get("file_data", b"")
        if self._is_html_content(file_data):
            self.logger.info(f"Detected HTML content in .{ext} file, processing as HTML table")
            return self._extract_html_as_excel(current_file)

        if ext in XLSX_FAMILY:
            return self._extract_xlsx(current_file, extract_metadata)
        elif ext == 'xls':
            return self._extract_xls(current_file, extract_metadata)
        else:
            raise ValueError(f"Unsupported Excel format: {ext}")

    def _spreadsheet_options(self) -> Dict[str, Any]:
        opts = {}
        try:
            opts = dict((self._config or {}).get("spreadsheet") or {})
        except Exception:
            opts = {}
        return {
            "include_hidden": bool(opts.get("include_hidden", True)),
            "streaming_threshold_bytes": int(opts.get("streaming_threshold_bytes", STREAMING_THRESHOLD_BYTES)),
        }

    def _visible_grid(self, grid: SheetGrid, include_hidden: bool) -> Optional[SheetGrid]:
        """숨김 옵션 적용. 시트를 통째로 뺄 때는 None."""
        if include_hidden:
            return grid
        if grid.hidden:
            return None
        grid.drop_hidden_cells()
        return grid

    def extract_text_fast(self, current_file: "CurrentFile") -> str:
        """
        Fast plain-text extraction for pre-scan.

        Reads cell values only. Skips: textbox, embedded images, charts, formatting.
        XLSX 는 openpyxl, XLS 는 xlrd 직접 사용.
        """
        file_path = current_file.get("file_path", "unknown")
        file_data = current_file.get("file_data", b"")
        ext = current_file.get("file_extension", os.path.splitext(file_path)[1]).lower().lstrip('.')
        self.logger.info(f"[Excel fast] Plain text extraction: {file_path} ext={ext}")

        try:
            import io as _io
            lines = []
            if ext in XLSX_FAMILY:
                from openpyxl import load_workbook
                wb = load_workbook(_io.BytesIO(file_data), read_only=True, data_only=True)
                for sheet in wb.worksheets:
                    for row in sheet.iter_rows(values_only=True):
                        row_vals = [str(v) for v in row if v is not None and str(v).strip()]
                        if row_vals:
                            lines.append("\t".join(row_vals))
                wb.close()
            elif ext == 'xls':
                import xlrd
                book = xlrd.open_workbook(file_contents=file_data)
                for sheet in book.sheets():
                    for row_idx in range(sheet.nrows):
                        row_vals = []
                        for col_idx in range(sheet.ncols):
                            v = sheet.cell_value(row_idx, col_idx)
                            if v is not None and str(v).strip():
                                row_vals.append(str(v))
                        if row_vals:
                            lines.append("\t".join(row_vals))
            else:
                # 알 수 없는 ext — fallback
                return self.extract_text(current_file, extract_metadata=False)
            return "\n".join(lines)
        except Exception as e:
            self.logger.warning(f"[Excel fast] Error on {file_path}: {e} — falling back to extract_text")
            return self.extract_text(current_file, extract_metadata=False)

    def _extract_xlsx(
        self,
        current_file: "CurrentFile",
        extract_metadata: bool = True
    ) -> str:
        """XLSX file processing."""
        file_path = current_file.get("file_path", "unknown")
        self.logger.info(f"XLSX processing: {file_path}")

        try:
            file_data = current_file.get("file_data", b"")
            opts = self._spreadsheet_options()
            # 큰 워크북은 전체 모드로 열지 않는다. 메모리가 시트 XML 의 약 12배 든다.
            streaming = is_large_workbook(file_data, opts["streaming_threshold_bytes"])
            if streaming:
                wb = load_streaming_workbook(file_data)
                self.logger.info(f"XLSX streaming mode (large workbook): {file_path}")
            else:
                # Step 1: Convert to Workbook using file_converter
                wb = self.file_converter.convert(file_data, extension='xlsx')
                # Step 2: Preprocess - may transform wb in the future
                preprocessed = self.preprocess(wb)
                wb = preprocessed.clean_content  # TRUE SOURCE

            try:
                grids = grids_from_workbook(wb, file_data, read_only=streaming)
                preload = self._preload_xlsx_data(current_file, wb, extract_metadata)
            finally:
                if streaming:
                    try:
                        wb.close()
                    except Exception:
                        pass

            result_parts = [preload["metadata_str"]] if preload["metadata_str"] else []
            processed_images: Set[str] = set()
            stats = {"charts": 0, "images": 0, "textboxes": 0, "sheet_images": 0}

            worksheets = [None] * len(grids) if streaming else list(wb.worksheets)
            for ws, grid in zip(worksheets, grids):
                grid = self._visible_grid(grid, opts["include_hidden"])
                if grid is None:
                    continue
                sheet_result = self._process_xlsx_sheet(
                    ws, grid, preload, processed_images, stats
                )
                result_parts.append(sheet_result)

            remaining = self._process_remaining_charts(
                preload["chart_data_list"], preload["chart_idx"], processed_images, stats
            )
            if remaining:
                result_parts.append(remaining)

            # 어느 시트에도 붙지 않은 이미지(스트리밍 모드, 또는 openpyxl 이 시트 이미지를 못 읽은 경우)는
            # 끝에 한 번만 낸다. 예전에는 시트마다 전체 이미지를 다시 붙였다.
            if stats["sheet_images"] == 0 and preload["images_data"]:
                for _name, img_data in preload["images_data"].items():
                    image_tag = self.format_image_processor.save_image(img_data)
                    if image_tag:
                        result_parts.append(f"\n{image_tag}\n")
                        stats["images"] += 1

            result = "".join(result_parts)
            self.logger.info(
                f"XLSX processing completed: {len(grids)} sheets, "
                f"{stats['charts']} charts, {stats['images']} images"
            )
            return result

        except Exception as e:
            self.logger.error(f"Error in XLSX processing: {e}")
            import traceback
            self.logger.debug(traceback.format_exc())
            raise

    def _extract_xls(
        self,
        current_file: "CurrentFile",
        extract_metadata: bool = True
    ) -> str:
        """XLS file processing."""
        file_path = current_file.get("file_path", "unknown")
        self.logger.info(f"XLS processing: {file_path}")

        try:
            # Step 1: Convert to Workbook using file_converter
            file_data = current_file.get("file_data", b"")
            wb = self.file_converter.convert(file_data, extension='xls')

            # Step 2: Preprocess - may transform wb in the future
            preprocessed = self.preprocess(wb)
            wb = preprocessed.clean_content  # TRUE SOURCE

            result_parts = []
            stats = {"images": 0, "textboxes": 0}

            if extract_metadata:
                xls_extractor = self._get_xls_metadata_extractor()
                metadata_str = xls_extractor.extract_and_format(wb)
                if metadata_str:
                    result_parts.append(metadata_str + "\n\n")

            # Extract images grouped by sheet using XLSImageProcessor
            xls_image_processor = self._create_xls_image_processor()
            sheet_names = [wb.sheet_by_index(i).name for i in range(wb.nsheets)]
            images_by_sheet = xls_image_processor.extract_images_by_sheet(file_path, sheet_names)

            # Extract textboxes/shape texts from XLS
            textboxes_by_sheet = extract_textboxes_from_xls(file_path, sheet_names)

            opts = self._spreadsheet_options()
            for sheet_idx in range(wb.nsheets):
                ws = wb.sheet_by_index(sheet_idx)
                grid = self._visible_grid(grid_from_xlrd(ws, wb), opts["include_hidden"])
                if grid is None:
                    continue
                sheet_tag = self.create_sheet_tag(ws.name)
                result_parts.append(f"\n{sheet_tag}\n")

                # Process tables/text blocks for this sheet
                result_parts.extend(f"\n{part}\n" for part in render_sheet(grid))

                # Process images for this sheet
                sheet_images = images_by_sheet.get(sheet_idx, [])
                for idx, img_data in enumerate(sheet_images):
                    if img_data:
                        image_tag = xls_image_processor.save_image(img_data)
                        if image_tag:
                            result_parts.append(f"\n{image_tag}\n")
                            stats["images"] += 1

                # Process textboxes/shape texts for this sheet
                sheet_textboxes = textboxes_by_sheet.get(ws.name, [])
                for tb in sheet_textboxes:
                    if tb:
                        result_parts.append(f"\n[Textbox] {tb}\n")
                        stats["textboxes"] += 1

            result = "".join(result_parts)
            self.logger.info(f"XLS processing completed: {wb.nsheets} sheets, {stats['images']} images, {stats['textboxes']} textboxes")
            return result

        except Exception as e:
            self.logger.error(f"Error in XLS processing: {e}")
            import traceback
            self.logger.debug(traceback.format_exc())
            raise

    @staticmethod
    def _is_html_content(file_data: bytes) -> bool:
        """
        Check if file content is actually HTML (not a real Excel binary/ZIP).

        Some systems (e.g., government websites) export HTML tables with
        .xls/.xlsx extensions. Excel can open these, but xlrd/openpyxl cannot.

        Note: Valid XLS always starts with OLE magic (\\xd0\\xcf\\x11\\xe0...),
        and valid XLSX always starts with ZIP magic (PK\\x03\\x04).
        Neither can ever start with '<html' or '<!doctype', so these checks
        are safe with zero false positives.

        Files starting with '<?xml' need extra verification because
        Excel 2003 XML Spreadsheet format also starts with '<?xml' but
        is NOT HTML — it contains <Workbook> instead of <html>.
        """
        if not file_data or len(file_data) < 20:
            return False
        # Check first 1024 bytes for HTML signatures (skip BOM if present)
        header = file_data[:1024].lstrip(b'\xef\xbb\xbf').lstrip()
        header_lower = header.lower()
        # Definitive HTML signatures
        if header_lower.startswith(b'<html') or header_lower.startswith(b'<!doctype'):
            return True
        # For <?xml, verify <html> tag exists (exclude Excel 2003 XML Spreadsheet)
        if header_lower.startswith(b'<?xml'):
            return b'<html' in header_lower
        return False

    def _extract_html_as_excel(
        self,
        current_file: "CurrentFile",
    ) -> str:
        """
        Process an HTML file disguised as Excel (.xls/.xlsx).

        Parses HTML tables using BeautifulSoup and converts them to
        the same output format as regular Excel processing.

        Handles two patterns:
        - Single table with <thead>/<tbody> (standard HTML)
        - Header-only table followed by body-only table (e.g. government websites)
          → these are merged into one logical table before output
        """
        from bs4 import BeautifulSoup

        file_path = current_file.get("file_path", "unknown")
        file_data = current_file.get("file_data", b"")
        self.logger.info(f"HTML-as-Excel processing: {file_path}")

        try:
            # Decode HTML content
            text = self._decode_html_bytes(file_data)
            soup = BeautifulSoup(text, 'html.parser')

            tables = soup.find_all('table')
            if not tables:
                # No tables found - extract plain text
                body = soup.find('body')
                plain_text = body.get_text(separator='\n', strip=True) if body else soup.get_text(separator='\n', strip=True)
                return plain_text

            # Merge consecutive split tables:
            # A "header-only" table (has <thead> or only <th> cells, no <td>)
            # followed by a "body-only" table (has <tbody> or only <td> cells, no <th>)
            # are treated as a single logical table.
            logical_tables = self._merge_split_tables(tables)

            # 엑셀로 열면 파일 이름의 시트 하나가 된다. 일반 엑셀과 같은 격자·렌더러로 낸다
            # (예전에는 병합이 있으면 청크 분할기가 모르는 <table> 을, 없으면 머리글 한 줄짜리 표를 냈다).
            sheet_name = os.path.splitext(os.path.basename(str(file_path)))[0] or "Sheet1"
            result_parts = [f"\n{self.create_sheet_tag(sheet_name)}\n"]
            output_count = 0

            for logical_table in logical_tables:
                all_rows = logical_table['header_rows'] + logical_table['body_rows']
                if not all_rows:
                    continue
                parts = render_sheet(grid_from_html_rows(sheet_name, all_rows))
                if not parts:
                    continue
                output_count += 1
                result_parts.extend(f"\n{part}\n" for part in parts)

            result = "".join(result_parts)
            self.logger.info(
                f"HTML-as-Excel processing completed: {output_count} tables extracted"
            )
            return result

        except Exception as e:
            self.logger.error(f"Error in HTML-as-Excel processing: {e}")
            import traceback
            self.logger.debug(traceback.format_exc())
            raise

    @staticmethod
    def _merge_split_tables(tables) -> List[Dict]:
        """
        Analyse a list of <table> tags and merge consecutive pairs where
        the first contains only header rows (<thead> / <th>) and the second
        contains only body rows (<tbody> / <td>).

        Returns a list of dicts with keys:
            'header_rows': list of <tr> BeautifulSoup tags
            'body_rows':   list of <tr> BeautifulSoup tags
        """
        def classify(table):
            """Return ('header', rows) | ('body', rows) | ('mixed', rows)."""
            all_tr = table.find_all('tr')
            if not all_tr:
                return 'empty', []
            has_th = bool(table.find('th'))
            has_td = bool(table.find('td'))
            thead_rows = [tr for tr in table.find_all('tr') if tr.find_parent('thead')]
            tbody_rows = [tr for tr in table.find_all('tr') if tr.find_parent('tbody')]
            # Header-only: has <thead> but no <tbody>, or only <th> cells
            if has_th and not has_td:
                return 'header_only', all_tr
            # Body-only: has <tbody> but no <thead>, or only <td> cells
            if has_td and not has_th:
                return 'body_only', all_tr
            # Mixed: single table with both <thead> and <tbody>
            return 'mixed', (thead_rows or all_tr[:1], tbody_rows or all_tr[1:])

        logical = []
        i = 0
        while i < len(tables):
            kind, rows = classify(tables[i])
            if kind == 'empty':
                i += 1
                continue
            if kind == 'header_only' and i + 1 < len(tables):
                next_kind, next_rows = classify(tables[i + 1])
                if next_kind == 'body_only':
                    logical.append({'header_rows': rows, 'body_rows': next_rows})
                    i += 2
                    continue
            if kind == 'mixed':
                header_rows, body_rows = rows
                logical.append({'header_rows': header_rows, 'body_rows': body_rows})
            elif kind == 'header_only':
                logical.append({'header_rows': rows, 'body_rows': []})
            else:  # body_only or unmerged
                logical.append({'header_rows': [], 'body_rows': rows})
            i += 1
        return logical

    @staticmethod
    def _decode_html_bytes(file_data: bytes) -> str:
        """Decode HTML bytes to string with encoding fallback."""
        for enc in ['utf-8', 'utf-8-sig', 'cp949', 'euc-kr', 'latin-1']:
            try:
                return file_data.decode(enc)
            except UnicodeDecodeError:
                continue
        return file_data.decode('utf-8', errors='replace')

    def _preload_xlsx_data(
        self, current_file: "CurrentFile", wb, extract_metadata: bool
    ) -> Dict[str, Any]:
        """Extract preprocessing data from XLSX file."""
        file_path = current_file.get("file_path", "unknown")
        file_stream = self.get_file_stream(current_file)

        result = {
            "metadata_str": "",
            "chart_data_list": [],  # ChartData instances from extractor
            "images_data": [],
            "textboxes_by_sheet": {},
            "chart_idx": 0,
        }

        if extract_metadata:
            result["metadata_str"] = self.extract_and_format_metadata(wb)
            if result["metadata_str"]:
                result["metadata_str"] += "\n\n"

        # Use ChartExtractor for chart extraction
        result["chart_data_list"] = self.chart_extractor.extract_all_from_file(file_stream)

        # Use format_image_processor directly for image extraction
        image_processor = self.format_image_processor
        if hasattr(image_processor, 'extract_images_from_xlsx'):
            result["images_data"] = image_processor.extract_images_from_xlsx(file_path)
        else:
            result["images_data"] = {}
        result["textboxes_by_sheet"] = extract_textboxes_from_xlsx(file_path)

        return result

    def _process_xlsx_sheet(
        self, ws, grid: SheetGrid, preload: Dict[str, Any],
        processed_images: Set[str], stats: Dict[str, int]
    ) -> str:
        """Process a single XLSX sheet (ws is None in streaming mode)."""
        sheet_name = grid.name
        sheet_tag = self.create_sheet_tag(sheet_name)
        parts = [f"\n{sheet_tag}\n"]

        parts.extend(f"\n{part}\n" for part in render_sheet(grid))

        # Chart processing using ChartExtractor
        if ws is not None and hasattr(ws, '_charts') and ws._charts:
            chart_data_list = preload["chart_data_list"]
            for chart in ws._charts:
                if preload["chart_idx"] < len(chart_data_list):
                    chart_data = chart_data_list[preload["chart_idx"]]
                    # chart_data is already ChartData instance, format it
                    chart_output = self._format_chart_data(chart_data)
                    if chart_output:
                        parts.append(f"\n{chart_output}\n")
                        stats["charts"] += 1
                    preload["chart_idx"] += 1

        # Image processing: 이 시트에 붙은 이미지만(통째 목록으로 되돌아가지 않게 빈 사전을 넘긴다)
        image_processor = self.format_image_processor
        if ws is not None and hasattr(image_processor, 'get_sheet_images'):
            sheet_images = image_processor.get_sheet_images(ws, {}, "")
        else:
            sheet_images = []
        for image_data, anchor in sheet_images:
            if image_data:
                image_tag = self.format_image_processor.save_image(image_data)
                if image_tag:
                    parts.append(f"\n{image_tag}\n")
                    stats["images"] += 1
                    stats["sheet_images"] += 1

        # Textbox processing
        textboxes = preload["textboxes_by_sheet"].get(sheet_name, [])
        for tb in textboxes:
            if tb:
                parts.append(f"\n[Textbox] {tb}\n")
                stats["textboxes"] += 1

        return "".join(parts)

    def _format_chart_data(self, chart_data) -> str:
        """Format ChartData using ChartProcessor."""
        from xgen_doc2chunk.core.functions.chart_extractor import ChartData

        if not isinstance(chart_data, ChartData):
            return ""

        if chart_data.has_data():
            return self.chart_processor.format_chart_data(
                chart_type=chart_data.chart_type,
                title=chart_data.title,
                categories=chart_data.categories,
                series=chart_data.series
            )
        else:
            return self.chart_processor.format_chart_fallback(
                chart_type=chart_data.chart_type,
                title=chart_data.title
            )

    def _process_remaining_charts(
        self, chart_data_list: List, chart_idx: int,
        processed_images: Set[str], stats: Dict[str, int]
    ) -> str:
        """Process remaining charts not associated with sheets."""
        parts = []
        while chart_idx < len(chart_data_list):
            chart_data = chart_data_list[chart_idx]
            chart_output = self._format_chart_data(chart_data)
            if chart_output:
                parts.append(f"\n{chart_output}\n")
                stats["charts"] += 1
            chart_idx += 1
        return "".join(parts)


__all__ = ["ExcelHandler"]
