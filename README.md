# xgen-doc2chunk

**xgen-doc2chunk** is a document processing library that converts raw documents into AI-understandable context. It analyzes, restructures, and normalizes content so that language models can reason over documents with higher accuracy and consistency.

> **Current Version: 0.2.26** — See [CHANGELOG.md](CHANGELOG.md) for release history.

## Features

- **Multi-format Support**: Process a wide variety of document formats including:
  - PDF (adaptive complexity-based processing, multi-column layout, table detection)
  - Microsoft Office: DOCX, DOC, PPTX, PPT, XLSX, XLS
  - Korean documents: HWP, HWPX (Hangul Word Processor — full support)
  - Text formats: TXT, MD, RTF, CSV, TSV, HTML
  - Image files: JPG, PNG, GIF, BMP, WebP (via OCR)
  - Code files: Python, JavaScript, TypeScript, and 20+ languages
  - Config files: JSON, YAML, TOML, INI, ENV, and more

- **Intelligent Text Extraction**: 
  - Preserves document structure (headings, paragraphs, lists)
  - Extracts tables as HTML with proper `rowspan`/`colspan` handling
  - Handles merged cells and complex table layouts
  - Extracts and processes inline images
  - Header/footer extraction for DOC, DOCX, HWPX
  - Chart and diagram extraction from Office documents
  - Reads ISO 29500 **Strict** OOXML (Word's "Strict Open XML Document") and
    repairs packages with missing or broken root relationships

- **OCR Integration**:
  - Pluggable OCR engine architecture
  - Supports **OpenAI**, **Anthropic**, **Google Gemini**, **AWS Bedrock**, and **vLLM** backends
  - Automatic OCR fallback for scanned documents or image-based PDFs
  - Standalone image file processing (JPG, PNG, etc.)
  - Custom image tag pattern support for OCR detection

- **Smart Chunking**:
  - Semantic text chunking with configurable size and overlap
  - Table-aware chunking that preserves table integrity (HTML & Markdown)
  - Page-based chunking with page number metadata
  - Protected regions for code blocks, tables, images, charts, and metadata
  - Small chunk merging to prevent table-title isolation
  - Nested table support in protected region detection
  - Position metadata (page number, line numbers, character offsets)

- **Metadata Extraction**:
  - Extracts document metadata (title, author, creation date, etc.)
  - Formats metadata in a structured, parseable format
  - Customizable metadata tag prefixes/suffixes

- **Storage Backends**:
  - Local file storage (default)
  - MinIO / S3 compatible cloud storage
  - Pluggable storage backend architecture

## Installation

```bash
pip install xgen-doc2chunk
```

Or using uv:

```bash
uv add xgen-doc2chunk
```

## Quick Start

### Basic Usage

```python
from xgen_doc2chunk import DocumentProcessor

# Create processor instance
processor = DocumentProcessor()

# Extract text from a document
text = processor.extract_text("document.pdf")
print(text)

# Extract text and chunk in one step
result = processor.extract_chunks(
    "document.pdf",
    chunk_size=1000,
    chunk_overlap=200
)

# Access chunks
for i, chunk in enumerate(result.chunks):
    print(f"Chunk {i + 1}: {chunk[:100]}...")

# Save chunks to markdown file
result.save_to_md("output/chunks.md")
```

### With OCR Processing

```python
from xgen_doc2chunk import DocumentProcessor
from xgen_doc2chunk.ocr.ocr_engine.openai_ocr import OpenAIOCR

# Initialize OCR engine
ocr_engine = OpenAIOCR(api_key="sk-...", model="gpt-4o")

# Create processor with OCR
processor = DocumentProcessor(ocr_engine=ocr_engine)

# Extract text with OCR processing enabled
text = processor.extract_text(
    "scanned_document.pdf",
    ocr_processing=True
)
```

### With Position Metadata

```python
from xgen_doc2chunk import DocumentProcessor

processor = DocumentProcessor()

result = processor.extract_chunks(
    "document.pdf",
    chunk_size=1000,
    include_position_metadata=True
)

# Access position metadata per chunk
if result.has_metadata:
    for chunk_data in result.chunks_with_metadata:
        print(f"Page {chunk_data['page_number']}, "
              f"Lines {chunk_data['line_start']}-{chunk_data['line_end']}: "
              f"{chunk_data['text'][:80]}...")
```

### Available OCR Engines

```python
from xgen_doc2chunk.ocr.ocr_engine.openai_ocr import OpenAIOCR
from xgen_doc2chunk.ocr.ocr_engine.anthropic_ocr import AnthropicOCR
from xgen_doc2chunk.ocr.ocr_engine.gemini_ocr import GeminiOCR
from xgen_doc2chunk.ocr.ocr_engine.bedrock_ocr import BedrockOCR
from xgen_doc2chunk.ocr.ocr_engine.vllm_ocr import VllmOCR

# OpenAI (recommended)
engine = OpenAIOCR(api_key="sk-...", model="gpt-4o")

# Anthropic Claude
engine = AnthropicOCR(api_key="sk-ant-...", model="claude-sonnet-4-20250514")

# Google Gemini
engine = GeminiOCR(api_key="...", model="gemini-2.0-flash")

# AWS Bedrock
engine = BedrockOCR(
    aws_access_key_id="AKIA...",
    aws_secret_access_key="...",
    aws_region="us-east-1",
    model="anthropic.claude-3-5-sonnet-20241022-v2:0"
)

# vLLM (self-hosted)
engine = VllmOCR(base_url="http://localhost:8000", model="Qwen/Qwen2-VL-7B-Instruct")
```

## Supported Formats

| Category | Extensions |
|----------|------------|
| Documents | `.pdf`, `.docx`, `.doc`, `.rtf`, `.pptx`, `.ppt`, `.hwp`, `.hwpx` |
| Spreadsheets | `.xlsx`, `.xls`, `.csv`, `.tsv` |
| Text | `.txt`, `.md`, `.markdown` |
| Web | `.html`, `.htm`, `.xhtml` |
| Images | `.jpg`, `.jpeg`, `.png`, `.gif`, `.bmp`, `.webp` |
| Code | `.py`, `.js`, `.ts`, `.jsx`, `.tsx`, `.java`, `.cpp`, `.c`, `.go`, `.rs`, `.cs`, `.swift`, `.kt`, `.rb`, `.php`, `.dart`, `.r`, `.scala`, `.sql`, `.vue`, `.svelte` |
| Config | `.json`, `.yaml`, `.yml`, `.xml`, `.toml`, `.ini`, `.cfg`, `.conf`, `.properties`, `.env` |
| Script | `.sh`, `.bat`, `.ps1`, `.zsh`, `.fish` |
| Log | `.log` |

## Architecture

```
xgen_doc2chunk/
├── core/
│   ├── document_processor.py       # Main entry point (DocumentProcessor, ChunkResult)
│   ├── processor/                  # Format-specific handlers
│   │   ├── base_handler.py         # Abstract base handler
│   │   ├── pdf_handler.py          # PDF processing (PyMuPDF + pdfplumber)
│   │   ├── docx_handler.py         # DOCX processing
│   │   ├── doc_handler.py          # DOC processing (auto-detects format)
│   │   ├── ppt_handler.py          # PowerPoint processing
│   │   ├── excel_handler.py        # Excel processing (XLSX/XLS)
│   │   ├── csv_handler.py          # CSV/TSV processing
│   │   ├── hwp_handler.py          # HWP (OLE) processing
│   │   ├── hwpx_handler.py         # HWPX (ZIP/XML) processing
│   │   ├── rtf_handler.py          # RTF processing
│   │   ├── text_handler.py         # Plain text / code processing
│   │   ├── html_reprocessor.py     # HTML document processing
│   │   ├── image_file_handler.py   # Standalone image processing (via OCR)
│   │   └── {format}_helper/        # Format-specific utilities
│   └── functions/
│       ├── img_processor.py        # Image handling & tag generation
│       ├── page_tag_processor.py   # Page/slide/sheet tag processing
│       ├── chart_extractor.py      # Chart data extraction
│       ├── chart_processor.py      # Chart formatting
│       ├── metadata_extractor.py   # Metadata extraction & formatting
│       ├── table_extractor.py      # Table data structures
│       ├── table_processor.py      # Table formatting (HTML/Markdown/Text)
│       ├── storage_backend.py      # Pluggable storage (Local, MinIO, S3)
│       ├── preprocessor.py         # File preprocessing
│       ├── file_converter.py       # File format conversion
│       └── utils.py                # General utilities
├── chunking/
│   ├── chunking.py                 # Main chunking API
│   ├── text_chunker.py             # Text-based chunking
│   ├── table_chunker.py            # Table-aware chunking (HTML & Markdown)
│   ├── page_chunker.py             # Page-based chunking
│   ├── sheet_processor.py          # Sheet/metadata processing
│   ├── protected_regions.py        # Protected region detection (nested tables)
│   └── constants.py                # Constants and patterns
└── ocr/
    ├── base.py                     # BaseOCR abstract class
    ├── ocr_processor.py            # OCR processing utilities
    └── ocr_engine/                 # OCR engine implementations
        ├── openai_ocr.py           # OpenAI GPT-4 Vision
        ├── anthropic_ocr.py        # Anthropic Claude Vision
        ├── gemini_ocr.py           # Google Gemini Vision
        ├── bedrock_ocr.py          # AWS Bedrock Vision
        └── vllm_ocr.py             # vLLM (self-hosted)
```

## Requirements

- Python 3.12+
- Required dependencies are automatically installed (see `pyproject.toml`)

### System Dependencies

For full functionality, you may need:

- **Tesseract OCR**: For local OCR fallback
- **LibreOffice**: For DOC/RTF conversion (optional)
- **Poppler**: For PDF image extraction

## Tag Customization

```python
processor = DocumentProcessor(
    # Image tag format (default: [Image:path])
    image_directory="output/images",
    image_tag_prefix="[Image:",
    image_tag_suffix="]",
    
    # Page tag format (default: [Page Number: N])
    page_tag_prefix="[Page Number: ",
    page_tag_suffix="]",
    
    # Slide tag format (default: [Slide Number: N])
    slide_tag_prefix="[Slide Number: ",
    slide_tag_suffix="]",
    
    # Chart tag format
    chart_tag_prefix="[chart]",
    chart_tag_suffix="[/chart]",
    
    # Metadata tag format
    metadata_tag_prefix="<Document-Metadata>",
    metadata_tag_suffix="</Document-Metadata>",
)
```

## Documentation

- [QUICKSTART.md](QUICKSTART.md) — Comprehensive guide with pipeline overview, OCR setup, and examples
- [CHANGELOG.md](CHANGELOG.md) — Release history
- [CONTRIBUTING.md](CONTRIBUTING.md) — Contribution guidelines

## License

Apache License 2.0 — see [LICENSE](LICENSE) for details.

## Contributing

Contributions are welcome! Please see [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines.

### Repositories

This project is maintained in two mirrored repositories that are kept in sync automatically
by `.github/workflows/sync.yml` (`main` and `deploy` branches):

- https://github.com/master0419/xgen_doc2chunk
- https://github.com/PlateerLab/xgen-doc2chunk

You can open a branch or PR against either one; merged changes propagate to the other.
