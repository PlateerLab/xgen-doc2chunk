from xgen_doc2chunk.chunking.chunking import create_chunks
from xgen_doc2chunk.chunking.table_chunker import (
    chunk_large_markdown_table,
    chunk_large_table,
    split_long_markdown_row,
)

META = "<Document-Metadata>\n  작성자: test\n</Document-Metadata>\n\n"


def _markdown_table(rows: int, preamble: str = "") -> str:
    lines = ([preamble] if preamble else []) + ["| 번호 | 이름 |", "| --- | --- |"]
    lines += [f"| {i} | 이름{i:04d} |" for i in range(rows)]
    return "\n".join(lines)


def test_markdown_chunks_respect_chunk_size():
    chunks = chunk_large_markdown_table(_markdown_table(400), 500, 0)
    assert len(chunks) > 1
    assert max(len(c) for c in chunks) <= 500


def test_markdown_preamble_repeated():
    chunks = chunk_large_markdown_table(_markdown_table(400, "[Table 2] 명단"), 500, 0, context_prefix="[Sheet: A]")
    assert all(c.startswith("[Sheet: A]\n[Table 2] 명단\n") for c in chunks)
    assert all("| 번호 | 이름 |" in c for c in chunks)


def test_html_chunks_respect_chunk_size():
    rows = "".join(f"<tr><td>{i}</td><td>이름{i:04d}</td></tr>" for i in range(300))
    html = f"<table border='1'><tr><th>번호</th><th>이름</th></tr>{rows}</table>"
    chunks = chunk_large_table(html, 600, 0)
    assert len(chunks) > 1
    assert max(len(c) for c in chunks) <= 600


def test_long_row_split_keeps_width():
    row = "| 7 | 제목 | " + ("가나다 " * 1000).strip() + " |"
    pieces = split_long_markdown_row(row, 500)
    assert len(pieces) > 1
    assert all(p.startswith("| 7 | 제목 | ") and p.count("|") == 4 for p in pieces)
    assert all(len(p) <= 500 for p in pieces)


def test_wide_row_is_not_split():
    row = "| " + " | ".join(str(i) for i in range(2000)) + " |"
    assert split_long_markdown_row(row, 500) == [row]


def test_page_number_of_chunk_starting_with_marker():
    pages = [f"[Page Number: {n}]\n" + f"{n}페이지 문장입니다. " * 80 for n in range(1, 5)]
    text = META + "\n".join(pages)
    chunks = create_chunks(text, chunk_size=500, chunk_overlap=50, file_extension="pdf",
                           include_position_metadata=True)
    checked = 0
    for c in chunks:
        body = c["text"].split("</Document-Metadata>")[-1].lstrip()
        if body.startswith("[Page Number:"):
            assert c["page_number"] == int(body.split(":")[1].split("]")[0])
            checked += 1
    assert checked >= 3


def test_sheet_page_numbers_follow_sheet_order():
    text = META + "[Sheet: 매출]\n" + _markdown_table(200) + "\n\n[Sheet: 비용]\n" + _markdown_table(5)
    chunks = create_chunks(text, chunk_size=800, chunk_overlap=0, file_extension="xlsx",
                           include_position_metadata=True)
    pages = {("매출" if "[Sheet: 매출]" in c["text"] else "비용"): set() for c in chunks}
    for c in chunks:
        pages["매출" if "[Sheet: 매출]" in c["text"] else "비용"].add(c["page_number"])
    assert pages == {"매출": {1}, "비용": {2}}
