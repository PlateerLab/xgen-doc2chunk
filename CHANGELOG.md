# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.4.1] - 2026-09-23

### Fixed
- **HWP text decoding** (`HwpRecord.get_text`), per the HWP 5.0 spec table 6:
  - TAB is an 8-word inline control. It was read as one word, so the tab's 7 data words
    came out as stray characters (Cyrillic, Arabic, Syriac, ... - up to 200 per real
    document).
  - Ordinary characters are decoded as UTF-16LE runs, so a surrogate pair (Hancom PUA
    symbols such as U+F02B1) stays one character. It was split into two lone surrogates
    (552 across 23 real documents), which break UTF-8 encoding downstream
    (`surrogates not allowed`).
  - Hyphen (24), non-breaking space (30) and fixed-width space (31) are now text
    (`▪수소` → `▪ 수소`).
  - Output otherwise unchanged; extended-control markers (`\x0b`) are placed as before.

### Removed
- **pyhwp dependency** (AGPL-3.0). It was declared from the first release but never
  imported: HWP is read by the package's own parser (`hwp_helper`, olefile + zlib).
  Extraction output is byte-identical with and without pyhwp installed (8 HWP files,
  including real press releases and reports).
- `uv.lock` regenerated; it still carried pymupdf and pyhwp from before 0.4.0.

## [0.4.0] - 2026-09-21

### Changed
- **PDF handling moved from PyMuPDF (AGPL-3.0) to xgen-pdf** (pdfium-based, Apache-2.0).
  All `fitz` consumers import `xgen_pdf as fitz`; the call surface is unchanged.
  End-to-end extraction on six real PDFs keeps plain-text similarity 0.97-1.00 with
  identical page and image counts. xgen-pdf ships as a GitHub Release wheel.

## [0.3.71] - 2026-09-18

### Fixed
- **DOCX**: A package whose zip entries use the host path separator
  (`word\\document.xml` instead of `word/document.xml`) failed with the same
  message 0.3.7 addressed -
  `no relationship of type '...officeDocument' in collection` - but for a different
  reason, so 0.3.7 did not help. The package wiring is intact; only the **names the
  parts are filed under** are wrong. APPNOTE 4.4.17.1 requires `/`, so such a file is
  malformed, but it is produced by real generators and Word opens it.
  - **Why it surfaced now**: CPython normalised `\` to `/` in `zipfile` up to 3.12.
    From 3.13 `namelist()` returns the name verbatim, so `_rels/.rels` is no longer
    found and every OPC reader reports a missing root relationship. The same file
    therefore reads on 3.12 and fails on 3.14 - the document was always malformed.
  - **Why 0.3.7 missed it**: `diagnose_ooxml_package()` looked parts up by their
    literal names, so `_find_main_part()` matched nothing and the diagnosis stopped at
    `no_main_part` (not repairable) before any of the 0.3.7 repairs were considered.
- `diagnose_ooxml_package()` now resolves every part by its **normalised** name, and
  reports the new kind `backslash_paths` so the symptom names the cause instead of
  being mistaken for `missing_root_rels`.
- `repair_ooxml_package()` rewrites the entry names on the way out. Part **bytes are
  untouched**; only the name each part is filed under changes.

### Notes
- A file that already reads today follows the identical path as in 0.3.7: repair is
  still attempted only after python-docx has raised.
- `backslash_paths` combines with the other repairs - a package that is both Strict and
  backslash-named is handled in one pass.

## [0.3.7] - 2026-09-16

### Fixed
- **DOCX**: A document that python-docx could not open was indexed as a chunk whose
  *content* was the error message itself -
  `[DOCX file processing failed: no relationship of type '...officeDocument' in collection]`.
  The upload reported success, and the error string then surfaced in search results.
  `DOCXHandler._extract_docx_simple_text()` now raises `RuntimeError` instead of
  returning that marker, matching `_extract_simple_text_fallback()`, which has always
  raised. The message carries a diagnosis rather than the raw exception.
  **Behaviour change**: a DOCX that cannot be extracted now fails the upload instead of
  silently indexing its own failure.
- **DOCX**: Three unrelated file states produced that one identical message, so the
  symptom never identified the cause. All three are now read correctly:
  - **ISO 29500 Strict** packages - Word's "Strict Open XML Document" save option,
    mandated by some public-sector document policies. Strict is a standard, not a
    malformed file; python-docx simply only knows the Transitional namespaces.
  - **Missing `_rels/.rels`** - an incomplete package from a third-party generator or
    a truncated transfer.
  - **Missing or dangling `officeDocument` relationship** - root relationships that do
    not point at a part present in the package.
  In each case `word/document.xml` is intact and only the package wiring is unreadable.
  Verified against all three, plus a main part typed as `application/xml`: every one now
  extracts byte-identical text and tables to the healthy original.

### Added
- `xgen_doc2chunk/core/functions/ooxml_repair.py`: rebuilds an unreadable OOXML package
  **in memory** - the input bytes are never modified - and hands the reader a package it
  can open.
  - `open_docx_document()` opens normally first and repairs only after that raised, so a
    file that reads correctly today follows the identical code path as in 0.3.6. Wired
    into `DOCXFileConverter.convert()`, `DOCXHandler.extract_text_fast()` and
    `DOCFileConverter._convert_docx()`.
  - `diagnose_ooxml_package()` / `describe_ooxml_failure()` name the cause
    (`strict`, `missing_root_rels`, `dangling_office_relationship`, `ole2`,
    `no_main_part`, ...) instead of leaving one message to mean five things.
  - `python -m xgen_doc2chunk.core.functions.ooxml_repair <file>` reports that diagnosis
    from the command line for support triage.
- Strict -> Transitional mapping covering 104 namespace and relationship-type pairs,
  taken verbatim from the Microsoft Open XML SDK (`OpenXmlNamespaceResolver`) rather
  than derived by hand. A missing pair does not fail loudly - it silently drops the
  content that uses it - so anything left unmapped after a rewrite is reported through
  `OoxmlDiagnosis.unmapped_strict_uris` and logged as a warning.
  Verified by round-tripping every XML part of a real document Transitional -> Strict ->
  Transitional: zero byte differences, zero unmapped URIs.
- Pre-release Office 2007 namespaces (`wordprocessingml/2006/3/main` and siblings) are
  mapped as well; they break the same reader in the same way.

### Compatibility
- Strict is a *subset* of Transitional, so mapping in this direction cannot lose
  constructs. The reverse direction is unsafe and is not attempted.
- Repair runs only after a normal open has raised, so no file that works in 0.3.6 takes
  a different path in 0.3.7.
- Relationship-type substitution is ordered longest-key-first. This is required, not
  cosmetic: `.../officeDocument/relationships` is a prefix of every relationship type,
  and `extendedProperties` / `customProperties` map to the hyphenated
  `extended-properties` / `custom-properties`, which prefix replacement would get wrong.
- Packages that are genuinely something else - an XLSX or PPTX renamed to `.docx`, a
  legacy OLE2 `.doc`, a ZIP with no main part - are not "repaired" into nonsense. They
  raise with a message naming what the file actually is.
- Only the DOCX read paths are wired up. `ooxml_repair` itself is format-agnostic and
  already recognises `xl/workbook.xml` and `ppt/presentation.xml`, but the XLSX and PPTX
  handlers are untouched in this release.

## [0.3.6] - 2026-09-10

### Fixed
- **PDF/DOCX/HWPX**: An image placed inside a table cell (typically a screenshot of
  text pasted into the cell) was dropped entirely. The cell extracted as empty **and**
  no `[Image:...]` tag was produced anywhere, so the OCR pass never saw the content -
  it was unrecoverable regardless of which OCR model was configured. Such a cell now
  receives an inline image tag at its own position, so OCR replaces it in place and
  the text lands in the correct `<td>`.
  - PDF: images were excluded by the table-bbox filter in
    `pdf_image_processor.extract_images_from_page()` (70% overlap rule) and the
    document-level size filters (`min_image_size=50`, `min_image_area=2500`) discarded
    cell-sized images.
  - DOCX: `DOCXTableExtractor._extract_cell_text()` read `w:t` only, ignoring
    `w:drawing` / `w:pict`.
  - HWPX: `hwpx_section._process_table()` did not pass the zip / BinItem context to
    the extractor, so `hp:pic` inside `hp:tc` could not be resolved.
- **OCR**: `process_text_with_ocr()`, `process_text_with_ocr_progress()` and
  `BaseOCR.process_text()` passed the model output straight to `re.sub()` as the
  *replacement* string, so backslashes and group references (`\1`, `\g<name>`) in the
  output were interpreted - corrupting the text or raising `re.error`. Replacement is
  now literal.

### Added
- `xgen_doc2chunk/ocr/table_cell_ocr.py`: table-cell aware tag replacement.
  `replace_image_tags()` resolves each distinct image once, keeps the original tag on
  failure, and flattens output that lands inside a `<td>`/`<th>` via
  `sanitize_ocr_text_for_table_cell()` - the default OCR prompt asks the VL model for
  HTML tables, which would otherwise break `chunking/table_parser.py`.
- `xgen_doc2chunk/core/functions/table_cell_image.py`: shared `CellImageConfig`,
  `merge_cell_image_tag()` and `match_images_to_cells()` used by the format extractors.
  `min_cell_coverage` (default 0.25) separates content from decoration: a screenshot
  pasted as cell content fills most of its cell, while a checkbox or bullet icon covers
  a few percent of it. Pixel size alone cannot tell them apart, so coverage is the
  gate - without it, 16x16 icons in a calendar table would each trigger a VL model call.
- `xgen_doc2chunk/core/processor/pdf_helpers/pdf_cell_image.py`: matches image
  placements to the table cell containing them.
- `DOCXTableExtractor.configure_images()` and `HWPXTableExtractor.configure_images()`
  enable cell image extraction; without them both extractors behave exactly as before.

### Compatibility
- Only cells that extract as **empty** receive a tag (`only_empty_cells`), so tables
  that already extract correctly follow the identical code path as 0.3.5.
- The document-level image pass is unchanged, including the PDF table-bbox exclusion,
  so an image outside a table is still emitted exactly as before and is never tagged
  twice. Dedup sets are shared with the handlers.
- OCR output *outside* a table cell is inserted verbatim as before; flattening applies
  only inside a cell.
- Charts inside DOCX cells are deliberately not resolved: chart content is consumed
  from a document-ordered queue and pulling from it out of order would misalign the
  remaining charts.
- HWP already routed cell content through the image-aware traversal callback, and
  XLSX/XLS already emitted image tags at sheet level, so neither format is changed.

## [0.3.5] - 2026-09-01

### Fixed
- **Excel**: A single cell (or a single-row / single-column block) was detected as a
  table, so one stray value occupied an entire chunk. Blocks smaller than 2x2 are now
  emitted as plain text, matching the minimum table size DOCX/HWPX/PDF already enforce.
  No values are dropped - a small block keeps its cell values as text.
- **Chunking**: `chunk_multi_sheet_content()` emitted one chunk per segment regardless
  of size, so `chunk_size` acted only as an upper bound and never as a fill target.
  Segments that fit are now packed into a shared chunk. Occupancy is measured on
  segment bodies, because the context prefix (metadata + sheet marker) is written once
  per chunk rather than once per segment. Oversized segments keep their previous
  handling, and the buffer is always flushed at a sheet boundary so segments from
  different sheets are never merged.
- **Chunking**: Long plain-text segments in multi-sheet content exceeded `chunk_size`
  by the context prefix length - the full `chunk_size` was passed to the text splitter
  before the prefix was prepended to every resulting chunk.
- **Chunking**: `extract_content_segments()` raised `NameError` when called with the
  default image pattern (`IMAGE_TAG_PATTERN` was referenced but never imported).

### Added
- **Excel**: `convert_xlsx_objects_to_blocks()` / `convert_xls_objects_to_blocks()`
  classify each detected object as `table` or `text`.
- **Excel**: `convert_xlsx_object_to_text()` / `convert_xls_object_to_text()` render a
  block smaller than the minimum table size as plain text.
- **Excel**: `LayoutRange.is_table_like()` plus `MIN_TABLE_ROWS` / `MIN_TABLE_COLS`
  constants in `excel_layout_detector`.

### Compatibility
- `convert_xlsx_objects_to_tables()` / `convert_xls_objects_to_tables()` keep their
  previous behavior (every detected object becomes a table) and remain exported.
- Non-Excel formats (PDF/DOCX/DOC/HWP/HWPX/PPTX) are unaffected; extraction and
  chunking output is byte-identical to 0.3.4.

## [0.3.4] - 2026-08-09

### Fixed
- **Chunking**: Rowspan blocks larger than the chunk limit produced a single
  oversized chunk that exceeded embedding model input limits. With
  `force_chunking` enabled, such blocks are now split at row boundaries
  (rows are never cut) and the spanning merged cells are re-issued into each
  continuation chunk with adjusted rowspan, so every chunk remains a
  self-contained, structurally valid table.
- **Chunking**: `force_chunking` was silently ignored for table-based files
  (CSV/TSV/XLSX/XLS); it is now propagated to the table chunking path.
  Default behavior (`force_chunking=False`) is unchanged.

### Added
- **Chunking**: `SpanningCell` dataclass and `extract_spanning_cells()` helper
  for rowspan-aware cell extraction with content.

## [0.2.26] - 2026-04-03

### Added
- **HWPX**: Extract text from shapes and improve section processing
- **HWPX**: Improve header/footer handling

### Changed
- Bump version to 0.2.26

## [0.2.25] - 2026-03-28

### Added
- **DOCX**: Clean field codes and improve run element extraction
- **DOC/DOCX**: Add header and footer extraction
- **Excel**: HTML content detection in Excel cell processing

### Changed
- **PDF**: Refine text extraction logic to exclude lines within table bounding boxes

## [0.2.24] - 2026-03-20

### Added
- **Chunking**: Implement small chunk merging to prevent table-title isolation
- **Chunking**: Allow backward merging when blocked by page boundaries

## [0.2.23] - 2026-03-15

### Improved
- **Excel**: Enhance merged cell handling in XLS and XLSX HTML conversion

## [0.2.22] - 2026-03-10

### Changed
- Update version retrieval mechanism in `__init__.py`

## [0.2.21] - 2026-03-05

### Changed
- Minor internal improvements and stabilization

## [0.2.20] - 2026-02-28

### Improved
- **Excel**: Enhance XLSX and XLS layout detection to consider cells with borders as valid

## [0.2.18] - 2026-02-22

### Changed
- **Excel**: Update HTML conversion to treat all cells as data cells without header distinction

## [0.2.17] - 2026-02-18

### Changed
- **Excel**: Remove textbox and image segment extraction from `sheet_processor` to prevent each image/textbox from occupying a separate chunk

### Added
- **Excel**: XLS and XLSX textbox extraction support
- **Excel**: Separate XLS and XLSX image handler refactoring

## [0.2.14] - 2026-02-12

### Fixed
- **Chunking**: Enhance `clean_chunks` to merge page-marker-only chunks with next chunk (solves skipped page numbers)

## [0.2.13] - 2026-02-08

### Added
- **Chunking**: Support for nested tables (tables within tables within tables) in protected region detection

## [0.2.12] - 2026-02-05

### Fixed
- **PDF**: Adjust Y gap threshold for table merging in `TableDetectionEngine` to prevent merging of separate tables

## [0.2.11] - 2026-02-02

### Changed
- **PDF**: Refactor import statements in `pdf_table_detection.py`

## [0.2.1] - 2026-01-30

### Fixed
- **PDF**: Enhance text extraction logic to handle table region extraction duplication problem

## [0.2.0] - 2026-01-28

### Changed
- Improve file extension handling in `DocumentProcessor`
- Major version bump: stabilization of core API

## [0.1.5x] - 2026-01-24 ~ 2026-01-27

### Added
- **PDF**: CJK compatibility handling and fragmented text reconstruction
- **Excel**: Table processing with context extraction and improved chunking logic (respects `chunk_size`)
- **Chunking**: Enhanced chunking logic for handling chunk size constraints

### Fixed
- **PDF**: Table quality validation criteria adjustment for paragraph text detection

## [0.1.4] - 2026-01-22

### Changed
- Refactor: Adjust validation criteria for paragraph text detection in `TableQualityValidator`
- Improve comments and documentation across processors (Korean → English)

## [0.1.2] - 2026-01-20

### Added
- **BedrockOCR**: AWS Bedrock Vision model support for OCR processing
  - Supports Claude 3.5 Sonnet and other Bedrock vision models
  - Full AWS credential configuration (access key, secret key, session token, region)
  - Configurable timeouts and retry settings
- **ImageFileHandler**: New handler for standalone image files (jpg, png, gif, bmp, webp)
  - Automatically uses OCR engine when available
  - Returns image tag format when OCR is not configured for later processing
- **PageTagProcessor**: Centralized page/slide/sheet tag processing system
  - Unified tag generation across all document handlers
  - Configurable tag prefixes and suffixes
- **Image pattern support for OCR**: Custom image tag patterns now passed to OCR engine
  - `ImageProcessor.get_pattern_string()` method for regex pattern generation
  - `BaseOCR.set_image_pattern()` and `set_image_pattern_from_string()` methods
  - OCR engines now recognize custom image tag formats

### Changed
- **DocumentProcessor**: OCR engine setter now invalidates handler registry for proper refresh
- **Handler registry**: ImageFileHandler automatically registered with OCR engine support
- **QUICKSTART.md**: Complete rewrite with comprehensive documentation
  - 3-stage processing pipeline documentation (File → Text → OCR → Chunks)
  - Detailed OCR configuration guide for all 5 engines
  - Tag customization examples (image, page, slide, sheet)
  - Complete API reference with all parameters

### Improved
- All Korean comments and docstrings in `img_processor.py` converted to English
- Enhanced OCR integration with custom pattern matching support
- Better separation of concerns with PageTagProcessor

## [0.1.0] - 2026-01-19

### Added
- Initial release of xgen_doc2chunk
- Multi-format document support (PDF, DOCX, DOC, XLSX, XLS, PPTX, PPT, HWP, HWPX)
- Intelligent text extraction with structure preservation
- Table detection and extraction with HTML formatting
- OCR integration (OpenAI, Anthropic, Google Gemini, vLLM)
- Smart chunking with semantic awareness
- Metadata extraction
- Support for 20+ code file formats
- Korean document support (HWP, HWPX)

### Features
- `DocumentProcessor` class for easy document processing
- Configurable chunk size and overlap
- Protected regions for code blocks
- Pluggable OCR engine architecture
- Automatic encoding detection for text files
- Chart and image extraction from Office documents

[0.2.26]: https://github.com/master0419/doc2chunk/compare/v0.2.25...v0.2.26
[0.2.25]: https://github.com/master0419/doc2chunk/compare/v0.2.24...v0.2.25
[0.2.24]: https://github.com/master0419/doc2chunk/compare/v0.2.23...v0.2.24
[0.2.23]: https://github.com/master0419/doc2chunk/compare/v0.2.22...v0.2.23
[0.2.22]: https://github.com/master0419/doc2chunk/compare/v0.2.21...v0.2.22
[0.2.21]: https://github.com/master0419/doc2chunk/compare/v0.2.20...v0.2.21
[0.2.20]: https://github.com/master0419/doc2chunk/compare/v0.2.18...v0.2.20
[0.2.18]: https://github.com/master0419/doc2chunk/compare/v0.2.17...v0.2.18
[0.2.17]: https://github.com/master0419/doc2chunk/compare/v0.2.14...v0.2.17
[0.2.14]: https://github.com/master0419/doc2chunk/compare/v0.2.13...v0.2.14
[0.2.13]: https://github.com/master0419/doc2chunk/compare/v0.2.12...v0.2.13
[0.2.12]: https://github.com/master0419/doc2chunk/compare/v0.2.11...v0.2.12
[0.2.11]: https://github.com/master0419/doc2chunk/compare/v0.2.1...v0.2.11
[0.2.1]: https://github.com/master0419/doc2chunk/compare/v0.2.0...v0.2.1
[0.2.0]: https://github.com/master0419/doc2chunk/compare/v0.1.5...v0.2.0
[0.1.5x]: https://github.com/master0419/doc2chunk/compare/v0.1.4...v0.1.5
[0.1.4]: https://github.com/master0419/doc2chunk/compare/v0.1.2...v0.1.4
[0.1.2]: https://github.com/master0419/doc2chunk/compare/v0.1.0...v0.1.2
[0.1.0]: https://github.com/master0419/doc2chunk/releases/tag/v0.1.0
