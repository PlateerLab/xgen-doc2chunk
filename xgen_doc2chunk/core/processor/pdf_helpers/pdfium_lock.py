# xgen_doc2chunk/core/processor/pdf_helpers/pdfium_lock.py
"""
Process-wide lock for pdfium.

PDFium (used through xgen_pdf / pypdfium2) is not thread-safe: two threads
extracting PDFs at the same time garble the text (e.g. "æîô WHÅçL…"), can leave
the process in a state where every later PDF fails with "broken document", and
can crash the whole process with a segmentation fault (reproduced with two PDFs
on four threads, 2026-10-02).

Every PDF extraction in this package takes this lock, so callers that share a
process (an agent's document tool, attachment readers, indexers) never overlap
inside pdfium, whichever entry point they use. It is re-entrant because a fast
extraction may fall back to the full one.
"""

import threading

PDFIUM_LOCK = threading.RLock()
