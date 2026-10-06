import sheet_fixtures as F
from conftest import extract

from xgen_doc2chunk.core.processor.csv_helper.csv_parser import parse_csv_content
from xgen_doc2chunk.core.processor.csv_helper.csv_table import convert_rows_to_markdown


def test_blank_cells_stay_blank(processor, tmp_path):
    text, _ = extract(processor, F.csv_blank_cells(tmp_path))
    assert "<table" not in text
    assert "|  | 최수아 | 개발 |" in text
    assert "| 1003 | 박지호 |  |" in text


def test_no_row_cap(processor, tmp_path):
    _, chunks = extract(processor, F.csv_many_rows(tmp_path))
    rows = {l.split("|")[1].strip() for c in chunks for l in c["text"].splitlines()
            if l.startswith("| ") and l.split("|")[1].strip().isdigit()}
    assert len(rows) == 3000


def test_parser_keeps_rows_and_columns_beyond_old_limits():
    content = "\n".join(",".join(str(c) for c in range(1100)) for _ in range(3)) + "\n"
    rows = parse_csv_content(content, ",")
    assert len(rows) == 3 and len(rows[0]) == 1100
    content = "a\n" + "\n".join(str(i) for i in range(100_050))
    assert len(parse_csv_content(content, ",")) == 100_051


def test_tsv(processor, tmp_path):
    text, _ = extract(processor, F.tsv(tmp_path))
    assert "| 코드 | 이름 |" in text
    assert "| B2 | 배 |" in text


def test_markdown_pads_and_uses_letters_without_header():
    md = convert_rows_to_markdown([["1", "2", "3"], ["4", "5"]], has_header=False)
    assert md.splitlines() == ["| A | B | C |", "| --- | --- | --- |", "| 1 | 2 | 3 |", "| 4 | 5 |  |"]


def test_markdown_escapes_pipes_and_newlines():
    md = convert_rows_to_markdown([["k", "v"], ["a|b", "줄\n바꿈"]], has_header=True)
    assert "| a\\|b | 줄 바꿈 |" in md
