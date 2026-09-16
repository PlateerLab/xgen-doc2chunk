# xgen_doc2chunk/core/functions/ooxml_repair.py
"""
OOXML package repair - open packages that a Transitional-only reader refuses.

python-docx, like most OOXML readers, understands only the **Transitional**
flavour of ISO/IEC 29500. Three different real-world file states make it fail
with one and the same message, which is why the symptom alone never identifies
the cause::

    no relationship of type
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
    in collection

1. **Strict OOXML** - produced by Word's "Strict Open XML Document" save option,
   which some public-sector document policies mandate. The payload is identical;
   only the namespaces differ (``http://purl.oclc.org/ooxml/...``).
2. **Missing root relationships** - ``_rels/.rels`` absent, typically after a
   third-party generator or a truncated transfer.
3. **Missing or dangling officeDocument relationship** - ``_rels/.rels`` exists
   but does not point at a part that is actually in the package.

In all three cases ``word/document.xml`` itself is intact; only the package
wiring around it is unreadable. This module rebuilds that wiring **in memory**
and hands the reader a package it can open. The file on disk is never modified.

Strict -> Transitional is the lossless direction: Strict is a *subset* of
Transitional, so every Strict construct has a Transitional equivalent. The
mapping table below is lifted verbatim from the Microsoft Open XML SDK
(``DocumentFormat.OpenXml.Features.OpenXmlNamespaceResolver``) rather than
guessed, because a missing pair does not fail loudly - it silently drops the
affected content. Anything still unmapped after a rewrite is reported through
``OoxmlDiagnosis.unmapped_strict_uris`` so the gap stays visible.

Repair is strictly a **fallback**: callers try a normal open first and only come
here after it raised, so a file that reads correctly today follows the exact
same code path as before.

Usage::

    from xgen_doc2chunk.core.functions.ooxml_repair import open_docx_document
    doc = open_docx_document(file_data)          # repairs only when needed

    from xgen_doc2chunk.core.functions.ooxml_repair import diagnose_ooxml_package
    print(diagnose_ooxml_package(file_data).message)

    $ python -m xgen_doc2chunk.core.functions.ooxml_repair broken.docx
"""
from __future__ import annotations

import logging
import re
import zipfile
from dataclasses import dataclass, field
from io import BytesIO
from typing import Any, BinaryIO, Dict, List, Optional, Sequence, Tuple
from xml.etree import ElementTree as ET

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Package constants
# ---------------------------------------------------------------------------

ZIP_MAGIC = b"PK\x03\x04"
#: Compound File Binary header - a pre-2007 .doc / .xls / .ppt.
OLE2_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"

CONTENT_TYPES_PART = "[Content_Types].xml"
ROOT_RELS_PART = "_rels/.rels"

#: OPC (ISO/IEC 29500 part 2) namespaces. These are **not** versioned by the
#: Strict/Transitional split, so they are identical in both flavours.
PACKAGE_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
PACKAGE_CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

OFFICE_DOCUMENT_RELTYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
)

WML_DOCUMENT_MAIN_CT = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
)
SML_WORKBOOK_MAIN_CT = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
)
PML_PRESENTATION_MAIN_CT = (
    "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
)

#: Main part candidates in probe order, as (part name, content type, kind).
#: ``word/document2.xml`` is the name some Word builds use after a macro-enabled
#: document is converted, so it is probed as well.
MAIN_PART_CANDIDATES: Tuple[Tuple[str, str, str], ...] = (
    ("word/document.xml", WML_DOCUMENT_MAIN_CT, "Word"),
    ("word/document2.xml", WML_DOCUMENT_MAIN_CT, "Word"),
    ("xl/workbook.xml", SML_WORKBOOK_MAIN_CT, "Excel"),
    ("ppt/presentation.xml", PML_PRESENTATION_MAIN_CT, "PowerPoint"),
)

#: Matches any leftover Strict vocabulary URI, used to detect unmapped pairs.
_STRICT_URI_RE = re.compile(rb"http://purl\.oclc\.org/ooxml/[A-Za-z0-9._/-]*")

# ---------------------------------------------------------------------------
# Strict -> Transitional tables (Microsoft Open XML SDK, verbatim)
# ---------------------------------------------------------------------------

#: Vocabulary namespaces, as they appear in xmlns declarations.
STRICT_NAMESPACES: Dict[bytes, bytes] = {
    b"http://purl.oclc.org/ooxml/descriptions/base":                      b"http://descriptions.openxmlformats.org/description/base",
    b"http://purl.oclc.org/ooxml/descriptions/full":                      b"http://descriptions.openxmlformats.org/description/full",
    b"http://purl.oclc.org/ooxml/drawingml/chart":                        b"http://schemas.openxmlformats.org/drawingml/2006/chart",
    b"http://purl.oclc.org/ooxml/drawingml/chartDrawing":                 b"http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
    b"http://purl.oclc.org/ooxml/drawingml/compatibility":                b"http://schemas.openxmlformats.org/drawingml/2006/compatibility",
    b"http://purl.oclc.org/ooxml/drawingml/diagram":                      b"http://schemas.openxmlformats.org/drawingml/2006/diagram",
    b"http://purl.oclc.org/ooxml/drawingml/lockedCanvas":                 b"http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas",
    b"http://purl.oclc.org/ooxml/drawingml/main":                         b"http://schemas.openxmlformats.org/drawingml/2006/main",
    b"http://purl.oclc.org/ooxml/drawingml/picture":                      b"http://schemas.openxmlformats.org/drawingml/2006/picture",
    b"http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing":           b"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    b"http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing":        b"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    b"http://purl.oclc.org/ooxml/officeDocument/bibliography":            b"http://schemas.openxmlformats.org/officeDocument/2006/bibliography",
    b"http://purl.oclc.org/ooxml/officeDocument/customProperties":        b"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
    b"http://purl.oclc.org/ooxml/officeDocument/customXml":               b"http://schemas.openxmlformats.org/officeDocument/2006/customXml",
    b"http://purl.oclc.org/ooxml/officeDocument/customXmlDataProps":      b"http://schemas.openxmlformats.org/officeDocument/2006/customXmlDataProps",
    b"http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes":          b"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
    b"http://purl.oclc.org/ooxml/officeDocument/extendedProperties":      b"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
    b"http://purl.oclc.org/ooxml/officeDocument/math":                    b"http://schemas.openxmlformats.org/officeDocument/2006/math",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships":           b"http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/customXml": b"http://schemas.openxmlformats.org/officeDocument/2006/customXml",
    b"http://purl.oclc.org/ooxml/officeDocument/sharedTypes":             b"http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes",
    b"http://purl.oclc.org/ooxml/presentationml/main":                    b"http://schemas.openxmlformats.org/presentationml/2006/main",
    b"http://purl.oclc.org/ooxml/schemaLibrary/main":                     b"http://schemas.openxmlformats.org/schemaLibrary/2006/main",
    b"http://purl.oclc.org/ooxml/spreadsheetml/main":                     b"http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    b"http://purl.oclc.org/ooxml/wordprocessingml/main":                  b"http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}

#: Relationship type values, as they appear in Type= inside *.rels parts.
STRICT_RELATIONSHIP_TYPES: Dict[bytes, bytes] = {
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/aFChunk":                b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/attachedTemplate":       b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/audio":                  b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/calcChain":              b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/chart":                  b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/chartUserShapes":        b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartUserShapes",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet":             b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/commentAuthors":         b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/comments":               b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/connections":            b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/control":                b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/control",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/customProperties":       b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/customProperty":         b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/customProperty",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/customXml":              b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/customXmlProps":         b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/diagramColors":          b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/diagramData":            b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/diagramLayout":          b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/diagramQuickStyle":      b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/dialogsheet":            b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/drawing":                b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/endnotes":               b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/extendedProperties":     b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/externalLink":           b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/externalLinkPath":       b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/font":                   b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/font",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/fontTable":              b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/footer":                 b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/footnotes":              b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/frame":                  b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/frame",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/glossaryDocument":       b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/handoutMaster":          b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/header":                 b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/htmlPubSaveAs":          b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/htmlPubSaveAs",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/hyperlink":              b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/image":                  b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/mailMergeHeaderSource":  b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/mailMergeHeaderSource",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/mailMergeRecipientData": b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/mailMergeRecipientData",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/mailMergeSource":        b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/mailMergeSource",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/metadata/thumbnail":     b"http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/movie":                  b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/movie",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/notesMaster":            b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/notesSlide":             b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/numbering":              b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument":         b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/oleObject":              b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/package":                b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/package",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/pivotCacheDefinition":   b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/pivotCacheRecords":      b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/pivotTable":             b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/presProps":              b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/printerSettings":        b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/queryTable":             b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/revisionHeaders":        b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionHeaders",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/revisionLog":            b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionLog",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/settings":               b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/sharedStrings":          b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/sheetMetadata":          b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/slide":                  b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/slideLayout":            b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/slideMaster":            b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/slideUpdateInfo":        b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideUpdateInfo",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/slideUpdateUrl":         b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideUpdateUrl",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/styles":                 b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/subDocument":            b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/subDocument",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/table":                  b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/tableSingleCells":       b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableSingleCells",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/tableStyles":            b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/tags":                   b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/theme":                  b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/themeOverride":          b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/themeOverride",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/transform":              b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/transform",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/usernames":              b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/usernames",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/video":                  b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/video",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/viewProps":              b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/volatileDependencies":   b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/volatileDependencies",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/webSettings":            b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet":              b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships/xmlMaps":                b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/xmlMaps",
}

#: Pre-release namespaces emitted by early Office 2007 builds. Unrelated to the
#: Strict split, but they break the same reader in the same way, and handling
#: them here costs nothing.
LEGACY_NAMESPACES: Dict[bytes, bytes] = {
    b"http://schemas.microsoft.com/office/word/2010/11/wordml":        b"http://schemas.microsoft.com/office/word/2012/wordml",
    b"http://schemas.openxmlformats.org/drawingml/2006/3/main":        b"http://schemas.openxmlformats.org/drawingml/2006/main",
    b"http://schemas.openxmlformats.org/presentationml/2006/3/main":   b"http://schemas.openxmlformats.org/presentationml/2006/main",
    b"http://schemas.openxmlformats.org/spreadsheetml/2006/5/main":    b"http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    b"http://schemas.openxmlformats.org/spreadsheetml/2006/7/main":    b"http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    b"http://schemas.openxmlformats.org/wordprocessingml/2006/3/main": b"http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    b"http://schemas.openxmlformats.org/wordprocessingml/2006/5/main": b"http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    b"http://schemas.openxmlformats.org/wordprocessingml/2006/6/main": b"http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}


def _ordered(table: Dict[bytes, bytes]) -> Tuple[Tuple[bytes, bytes], ...]:
    """Longest key first.

    Mandatory, not cosmetic: ``.../officeDocument/relationships`` is a prefix of
    every relationship type, and two of those types are **not** formed by simple
    concatenation (``extendedProperties`` -> ``extended-properties``,
    ``customProperties`` -> ``custom-properties``). Replacing the short key
    first would produce a plausible-looking but wrong URI for those two.
    """
    return tuple(sorted(table.items(), key=lambda kv: -len(kv[0])))


# One Strict URI has two correct answers depending on where it appears:
# .../officeDocument/relationships/customXml is a relationship type inside a
# .rels part but a namespace anywhere else (a known defect in the ISO text that
# the SDK works around the same way). Hence two tables rather than one.
_XML_TABLE = _ordered({**STRICT_RELATIONSHIP_TYPES, **STRICT_NAMESPACES, **LEGACY_NAMESPACES})
_RELS_TABLE = _ordered({**STRICT_NAMESPACES, **STRICT_RELATIONSHIP_TYPES, **LEGACY_NAMESPACES})


# ---------------------------------------------------------------------------
# Diagnosis
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class OoxmlDiagnosis:
    """Why a package could not be opened, and whether repair can help."""

    #: Machine-readable cause: empty, ole2, not_zip, bad_zip, no_main_part,
    #: strict, missing_root_rels, missing_office_relationship,
    #: dangling_office_relationship, ok.
    kind: str
    #: Human-readable summary, safe to put in a log line or an error message.
    message: str
    #: Whether repair_ooxml_package() has anything to work with.
    repairable: bool
    #: Word / Excel / PowerPoint, or None when undetermined.
    package_kind: Optional[str] = None
    #: Main part name inside the package, e.g. word/document.xml.
    main_part: Optional[str] = None
    #: Strict URIs with no entry in the mapping tables. Non-empty means the
    #: corresponding content would be dropped - report these rather than hide
    #: them.
    unmapped_strict_uris: Tuple[str, ...] = field(default_factory=tuple)


def _localname(tag: Any) -> str:
    text = str(tag)
    return text.rsplit("}", 1)[-1] if "}" in text else text


def _normalize_part_name(value: Optional[str]) -> Optional[str]:
    """Turn an OPC target or part name into a plain zip entry name."""
    if not value:
        return None
    name = value.replace("\\", "/").strip()
    while name.startswith("./"):
        name = name[2:]
    return name[1:] if name.startswith("/") else name


def _find_main_part(names: Sequence[str]) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """Return (part name, content type, kind) of the first main part found."""
    lookup = set(names)
    for part, content_type, kind in MAIN_PART_CANDIDATES:
        if part in lookup:
            return part, content_type, kind
    return None, None, None


def _apply_table(data: bytes, table: Sequence[Tuple[bytes, bytes]]) -> Tuple[bytes, int]:
    replacements = 0
    for old, new in table:
        if old in data:
            replacements += data.count(old)
            data = data.replace(old, new)
    return data, replacements


def normalize_ooxml_xml(data: bytes, *, relationship_part: bool = False) -> Tuple[bytes, int]:
    """Rewrite Strict and pre-release namespaces to their Transitional form.

    Returns the rewritten bytes and the number of substitutions made.
    """
    return _apply_table(data, _RELS_TABLE if relationship_part else _XML_TABLE)


def _unmapped_strict_uris(entries: Sequence[Tuple[str, bytes]]) -> Tuple[str, ...]:
    """Strict URIs that survive a rewrite, i.e. that the tables do not cover."""
    found = set()
    for name, data in entries:
        rewritten, _ = normalize_ooxml_xml(data, relationship_part=name.endswith(".rels"))
        found.update(_STRICT_URI_RE.findall(rewritten))
    return tuple(sorted(uri.decode("ascii", "replace") for uri in found))


def _parse_xml(data: Optional[bytes]) -> Optional[ET.Element]:
    if not data:
        return None
    try:
        return ET.fromstring(data)
    except ET.ParseError:
        return None


def _office_document_targets(rels_data: bytes) -> Tuple[str, ...]:
    """Resolved targets of every officeDocument relationship in a rels part."""
    root = _parse_xml(rels_data)
    if root is None:
        return ()
    targets: List[str] = []
    for child in root:
        if _localname(child.tag) != "Relationship":
            continue
        if child.get("TargetMode") == "External":
            continue
        reltype = (child.get("Type") or "").strip()
        if not reltype.endswith("/officeDocument"):
            continue
        target = _normalize_part_name(child.get("Target"))
        if target:
            targets.append(target)
    return tuple(targets)


def diagnose_ooxml_package(file_data: bytes) -> OoxmlDiagnosis:
    """Classify why an OOXML package is unreadable.

    Never raises; an undiagnosable input simply comes back as not repairable.
    """
    if not file_data or len(file_data) < 4:
        return OoxmlDiagnosis("empty", "File is empty or truncated.", False)

    if not file_data.startswith(ZIP_MAGIC):
        if file_data.startswith(OLE2_MAGIC):
            return OoxmlDiagnosis(
                "ole2",
                "Legacy binary Office document (.doc/.xls/.ppt) carrying a .docx name.",
                False,
            )
        return OoxmlDiagnosis("not_zip", "Not a ZIP archive, so not an OOXML package.", False)

    try:
        with zipfile.ZipFile(BytesIO(file_data)) as archive:
            names = archive.namelist()
            entries = [
                (name, archive.read(name)) for name in names if name.endswith((".xml", ".rels"))
            ]
    except (zipfile.BadZipFile, zipfile.LargeZipFile, OSError, RuntimeError, EOFError) as exc:
        return OoxmlDiagnosis("bad_zip", f"ZIP header present but unreadable: {exc}", False)

    main_part, _, package_kind = _find_main_part(names)
    if main_part is None:
        return OoxmlDiagnosis(
            "no_main_part",
            "ZIP contains no OOXML main part "
            "(word/document.xml, xl/workbook.xml or ppt/presentation.xml).",
            False,
        )

    unmapped = _unmapped_strict_uris(entries)
    is_strict = any(_STRICT_URI_RE.search(data) for _, data in entries)
    part_data = dict(entries)

    def result(kind: str, message: str) -> OoxmlDiagnosis:
        return OoxmlDiagnosis(kind, message, True, package_kind, main_part, unmapped)

    if part_data.get(ROOT_RELS_PART) is None:
        return result(
            "missing_root_rels",
            f"Root relationship part {ROOT_RELS_PART} is missing; the package is incomplete.",
        )

    if is_strict:
        return result(
            "strict",
            "ISO 29500 Strict package (Word's 'Strict Open XML Document'). "
            "Readable once the namespaces are mapped to Transitional.",
        )

    targets = _office_document_targets(part_data[ROOT_RELS_PART])
    if not targets:
        return result(
            "missing_office_relationship",
            f"{ROOT_RELS_PART} declares no officeDocument relationship.",
        )
    if not any(target in set(names) for target in targets):
        return result(
            "dangling_office_relationship",
            f"officeDocument relationship points at {targets[0]!r}, "
            "which is not in the package.",
        )

    return OoxmlDiagnosis(
        "ok",
        "Package wiring looks correct; the failure is elsewhere in the document.",
        True,
        package_kind,
        main_part,
        unmapped,
    )


def describe_ooxml_failure(file_data: bytes) -> str:
    """One-line explanation suited to an error message shown to an operator."""
    diagnosis = diagnose_ooxml_package(file_data)
    if diagnosis.package_kind and diagnosis.package_kind != "Word":
        return (
            f"Package is {diagnosis.package_kind}, not Word "
            f"(main part is {diagnosis.main_part}); the .docx name is misleading."
        )
    return diagnosis.message


# ---------------------------------------------------------------------------
# Rewriting
# ---------------------------------------------------------------------------


def _escape_attr(value: str) -> str:
    return (
        value.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def _serialize(
    root_tag: str, namespace: str, children: Sequence[Tuple[str, Dict[str, str]]]
) -> bytes:
    """Emit a flat OPC part by hand.

    ElementTree parses here but never writes: register_namespace() mutates
    module-global state, and these parts are flat enough that building the text
    directly is both shorter and free of surprises.
    """
    out = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<{root_tag} xmlns="{namespace}">',
    ]
    for tag, attrs in children:
        rendered = " ".join(
            f'{key}="{_escape_attr(value)}"' for key, value in attrs.items() if value is not None
        )
        out.append(f"<{tag} {rendered}/>" if rendered else f"<{tag}/>")
    out.append(f"</{root_tag}>")
    return "".join(out).encode("utf-8")


def _next_relationship_id(used: set) -> str:
    index = 1
    while f"rId{index}" in used:
        index += 1
    return f"rId{index}"


def _rebuild_root_rels(rels_data: Optional[bytes], main_part: str, part_names: set) -> bytes:
    """Return a root rels part that definitely points at ``main_part``.

    Existing relationships are preserved verbatim - dropping them would cost the
    document its styles, images and numbering. Only the officeDocument
    relationship is added or repointed, and the Relationships element itself is
    re-emitted in the OPC namespace in case the original used a wrong one.
    """
    root = _parse_xml(rels_data)
    children: List[Tuple[str, Dict[str, str]]] = []
    used_ids: set = set()
    has_office = False

    if root is not None and _localname(root.tag) == "Relationships":
        for child in root:
            if _localname(child.tag) != "Relationship":
                continue
            attrs = dict(child.attrib)
            reltype = (attrs.get("Type") or "").strip()
            if reltype.endswith("/officeDocument") and attrs.get("TargetMode") != "External":
                if _normalize_part_name(attrs.get("Target")) not in part_names:
                    attrs["Target"] = main_part
                attrs["Type"] = OFFICE_DOCUMENT_RELTYPE
                has_office = True
            if attrs.get("Id"):
                used_ids.add(attrs["Id"])
            children.append(("Relationship", attrs))

    if not has_office:
        children.append(
            (
                "Relationship",
                {
                    "Id": _next_relationship_id(used_ids),
                    "Type": OFFICE_DOCUMENT_RELTYPE,
                    "Target": main_part,
                },
            )
        )

    return _serialize("Relationships", PACKAGE_RELS_NS, children)


def _rebuild_content_types(
    content_types_data: Optional[bytes], main_part: str, content_type: str
) -> Optional[bytes]:
    """Ensure ``main_part`` carries ``content_type``; return None if already fine.

    python-docx compares the main part's content type exactly, so a package that
    types it as plain application/xml - or as the macro-enabled variant - is
    rejected even once the relationships are correct.
    """
    extension = main_part.rsplit(".", 1)[-1].lower() if "." in main_part else ""
    root = _parse_xml(content_types_data)

    children: List[Tuple[str, Dict[str, str]]] = []
    effective: Optional[str] = None
    override_seen = False
    seen_extensions = set()

    if root is not None and _localname(root.tag) == "Types":
        for child in root:
            tag = _localname(child.tag)
            if tag not in ("Default", "Override"):
                continue
            attrs = dict(child.attrib)
            if tag == "Default":
                child_extension = (attrs.get("Extension") or "").lower()
                seen_extensions.add(child_extension)
                if child_extension == extension and effective is None:
                    effective = attrs.get("ContentType")
            elif _normalize_part_name(attrs.get("PartName")) == main_part:
                override_seen = True
                effective = attrs.get("ContentType")
                attrs["ContentType"] = content_type
            children.append((tag, attrs))

    if effective == content_type:
        return None

    if "xml" not in seen_extensions:
        children.insert(0, ("Default", {"Extension": "xml", "ContentType": "application/xml"}))
    if "rels" not in seen_extensions:
        children.insert(
            0,
            (
                "Default",
                {
                    "Extension": "rels",
                    "ContentType": "application/vnd.openxmlformats-package.relationships+xml",
                },
            ),
        )
    if not override_seen:
        children.append(("Override", {"PartName": f"/{main_part}", "ContentType": content_type}))

    return _serialize("Types", PACKAGE_CONTENT_TYPES_NS, children)


def _clone_zipinfo(info: zipfile.ZipInfo) -> zipfile.ZipInfo:
    """Copy the fields worth keeping, leaving sizes and CRC for writestr()."""
    clone = zipfile.ZipInfo(info.filename, info.date_time)
    clone.compress_type = info.compress_type
    clone.external_attr = info.external_attr
    clone.internal_attr = info.internal_attr
    clone.create_system = info.create_system
    return clone


def repair_ooxml_package(file_data: bytes) -> Optional[bytes]:
    """Return a readable copy of ``file_data``, or None if nothing can be done.

    Rewrites Strict namespaces, restores the root relationship to the main part
    and corrects the main part's content type. The input bytes are not modified.
    """
    try:
        with zipfile.ZipFile(BytesIO(file_data)) as source:
            names = source.namelist()
            main_part, content_type, _ = _find_main_part(names)
            if main_part is None or content_type is None:
                return None
            payload = [(info, source.read(info.filename)) for info in source.infolist()]
    except (zipfile.BadZipFile, zipfile.LargeZipFile, OSError, RuntimeError, EOFError) as exc:
        logger.debug("OOXML repair: package is not readable as a ZIP (%s)", exc)
        return None

    rewritten: List[Tuple[zipfile.ZipInfo, bytes]] = []
    substitutions = 0
    for info, data in payload:
        if info.filename.endswith((".xml", ".rels")):
            data, count = normalize_ooxml_xml(
                data, relationship_part=info.filename.endswith(".rels")
            )
            substitutions += count
        rewritten.append((info, data))

    by_name = {info.filename: data for info, data in rewritten}
    new_rels = _rebuild_root_rels(by_name.get(ROOT_RELS_PART), main_part, set(names))
    new_content_types = _rebuild_content_types(
        by_name.get(CONTENT_TYPES_PART), main_part, content_type
    )

    buffer = BytesIO()
    written = set()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as target:
        for info, data in rewritten:
            if info.filename == ROOT_RELS_PART:
                data = new_rels
            elif info.filename == CONTENT_TYPES_PART and new_content_types is not None:
                data = new_content_types
            written.add(info.filename)
            target.writestr(_clone_zipinfo(info), data)

        if ROOT_RELS_PART not in written:
            target.writestr(ROOT_RELS_PART, new_rels)
        if CONTENT_TYPES_PART not in written and new_content_types is not None:
            target.writestr(CONTENT_TYPES_PART, new_content_types)

    logger.debug(
        "OOXML repair: main part %s, %d namespace substitutions", main_part, substitutions
    )
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# Reader entry point
# ---------------------------------------------------------------------------


def open_docx_document(
    file_data: Optional[bytes] = None, file_stream: Optional[BinaryIO] = None
) -> Any:
    """Open a WordprocessingML document, repairing the package only if needed.

    A file that python-docx already accepts takes the identical path it took
    before this function existed; repair is attempted solely after a failure.

    Raises:
        The original python-docx exception when the package cannot be repaired
        or when repair does not help.
    """
    from docx import Document  # local: keeps this module importable without python-docx

    if file_stream is not None:
        file_stream.seek(0)
        raw = file_stream.read()
    else:
        raw = file_data or b""

    try:
        return Document(BytesIO(raw))
    except Exception as original:
        diagnosis = diagnose_ooxml_package(raw)
        if not diagnosis.repairable or diagnosis.package_kind != "Word":
            logger.warning("DOCX cannot be repaired (%s): %s", diagnosis.kind, diagnosis.message)
            raise

        repaired = repair_ooxml_package(raw)
        if repaired is None:
            logger.warning(
                "DOCX repair produced nothing (%s): %s", diagnosis.kind, diagnosis.message
            )
            raise

        try:
            document = Document(BytesIO(repaired))
        except Exception:
            logger.warning("DOCX repair did not help (%s): %s", diagnosis.kind, diagnosis.message)
            raise original from None

        logger.warning("DOCX opened after repair (%s): %s", diagnosis.kind, diagnosis.message)
        if diagnosis.unmapped_strict_uris:
            logger.warning(
                "OOXML repair: %d Strict namespace(s) have no Transitional mapping, "
                "content using them is dropped: %s",
                len(diagnosis.unmapped_strict_uris),
                ", ".join(diagnosis.unmapped_strict_uris),
            )
        return document


__all__ = [
    "OoxmlDiagnosis",
    "STRICT_NAMESPACES",
    "STRICT_RELATIONSHIP_TYPES",
    "LEGACY_NAMESPACES",
    "MAIN_PART_CANDIDATES",
    "diagnose_ooxml_package",
    "describe_ooxml_failure",
    "normalize_ooxml_xml",
    "repair_ooxml_package",
    "open_docx_document",
]


def _main(argv: Sequence[str]) -> int:
    """python -m xgen_doc2chunk.core.functions.ooxml_repair <file> [...]"""
    if not argv:
        print(__doc__)
        return 2
    for path in argv:
        with open(path, "rb") as handle:
            data = handle.read()
        diagnosis = diagnose_ooxml_package(data)
        print(f"file       : {path}")
        print(f"size       : {len(data):,} bytes")
        print(f"kind       : {diagnosis.kind}")
        print(f"package    : {diagnosis.package_kind or '-'} ({diagnosis.main_part or '-'})")
        print(f"repairable : {'yes' if diagnosis.repairable else 'no'}")
        print(f"detail     : {describe_ooxml_failure(data)}")
        if diagnosis.unmapped_strict_uris:
            print(f"unmapped   : {', '.join(diagnosis.unmapped_strict_uris)}")
        if diagnosis.repairable and diagnosis.package_kind == "Word":
            try:
                document = open_docx_document(data)
                print(
                    f"result     : opens - {len(document.paragraphs)} paragraphs, "
                    f"{len(document.tables)} tables"
                )
            except Exception as exc:  # pragma: no cover - diagnostic path
                print(f"result     : still fails - {type(exc).__name__}: {exc}")
        print()
    return 0


if __name__ == "__main__":  # pragma: no cover
    import sys

    raise SystemExit(_main(sys.argv[1:]))
