"""
시트 격자(SheetGrid)를 표·글 덩어리로 나누고 청킹하기 좋은 글자로 그린다.

덩어리 나누기
    1. 값(또는 테두리·병합)이 있는 셀을 상하좌우로 이어 덩어리를 만든다(행 구간 단위로 빠르게).
    2. 겹치거나 안에 들어간 덩어리는 합친다. 예전에는 빈 칸이 많은 표(근무표 O 표시 등)에서
       표 안의 떨어진 셀이 따로 잡혀 같은 값이 한 번 더 나왔다.
    3. 빈 열 한 줄을 사이에 둔 덩어리가 같은 행 범위면 한 표다(중간 빈 열로 표가 둘로 갈라지던 문제).
    4. 빈 행 한 줄을 사이에 둔 덩어리가 같은 열 범위이고 아래 덩어리 첫 행이 위 데이터와 같은
       모양이면 이어지는 표다(빈 행으로 그룹을 나눈 표에서 둘째 그룹이 데이터 행을 머리글로 쓰던 문제).

그리기
    - 한 행 또는 한 열짜리 덩어리는 글(표 문법 없이 값만).
    - 두 열이 빈 열을 사이에 두고 있으면 양식(라벨: 값)으로 그린다.
    - 그 밖은 마크다운 표. 위쪽의 제목 행(한 칸만 있는 행·표 너비로 병합된 행)은 표 제목으로,
      여러 줄 병합 머리글은 한 줄로 펴서(``1월 목표``) 쓴다. 청크로 나뉘어도 모든 조각에
      제목과 머리글이 실리게 하려는 것이다(청커가 머리글을 반복한다).
    - 데이터 영역의 세로 병합은 아래 행에도 값을 채운다(행마다 뜻이 완결되게).
    - 머리글이 없는 표는 열 문자(A, B, …)를 머리글로 쓴다. 데이터 행을 머리글로 쓰지 않는다.
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Set, Tuple

from xgen_doc2chunk.core.processor.excel_helper.sheet_grid import SheetGrid

__all__ = ["SheetBlock", "detect_blocks", "render_sheet"]

#: 이보다 넓은 병합 영역은 점유 판단에서 영역 전체를 채우지 않는다(병적인 파일 방어).
_MAX_MERGE_AREA = 200_000
_MAX_CAPTION_ROWS = 2
_MAX_HEADER_ROWS = 4


#: 표가 아닌 한 줄짜리 덩어리의 칸을 잇는 구분자.
TEXT_JOINER = " · "


def column_letter(col: int) -> str:
    out = ""
    while col > 0:
        col, rem = divmod(col - 1, 26)
        out = chr(65 + rem) + out
    return out


@dataclass
class Box:
    r1: int
    c1: int
    r2: int
    c2: int

    def overlaps(self, o: "Box") -> bool:
        return not (self.r2 < o.r1 or o.r2 < self.r1 or self.c2 < o.c1 or o.c2 < self.c1)

    def union(self, o: "Box") -> "Box":
        return Box(min(self.r1, o.r1), min(self.c1, o.c1), max(self.r2, o.r2), max(self.c2, o.c2))


@dataclass
class SheetBlock:
    box: Box
    skip_rows: Set[int] = field(default_factory=set)  # 이어 붙이며 버린 반복 머리글 행


class _UF:
    def __init__(self, n: int):
        self.p = list(range(n))

    def find(self, x: int) -> int:
        while self.p[x] != x:
            self.p[x] = self.p[self.p[x]]
            x = self.p[x]
        return x

    def union(self, a: int, b: int) -> None:
        ra, rb = self.find(a), self.find(b)
        if ra != rb:
            self.p[rb] = ra


# ──────────────────────────────────────────────────────────────────────────────
# 덩어리 나누기
# ──────────────────────────────────────────────────────────────────────────────

def _occupied_by_row(grid: SheetGrid) -> Dict[int, List[int]]:
    occ: Dict[int, Set[int]] = {}
    for r, c in grid.cells:
        occ.setdefault(r, set()).add(c)
    for r, c in grid.bordered:
        occ.setdefault(r, set()).add(c)
    for r1, c1, r2, c2 in grid.merges:
        if (r1, c1) not in grid.cells:
            continue
        if (r2 - r1 + 1) * (c2 - c1 + 1) > _MAX_MERGE_AREA:
            continue
        for r in range(r1, r2 + 1):
            occ.setdefault(r, set()).update(range(c1, c2 + 1))
    return {r: sorted(cs) for r, cs in occ.items()}


def _components(occ: Dict[int, List[int]]) -> List[Box]:
    """행마다 연속 구간(run)을 만들고, 윗행 구간과 열이 겹치면 묶는다."""
    runs: List[Tuple[int, int, int]] = []  # (row, c1, c2)
    row_runs: Dict[int, List[int]] = {}
    for r in sorted(occ):
        cols = occ[r]
        start = prev = cols[0]
        for c in cols[1:]:
            if c == prev + 1:
                prev = c
                continue
            row_runs.setdefault(r, []).append(len(runs))
            runs.append((r, start, prev))
            start = prev = c
        row_runs.setdefault(r, []).append(len(runs))
        runs.append((r, start, prev))
    uf = _UF(len(runs))
    for r, ids in row_runs.items():
        above = row_runs.get(r - 1)
        if not above:
            continue
        i = j = 0
        while i < len(ids) and j < len(above):
            _, a1, a2 = runs[ids[i]]
            _, b1, b2 = runs[above[j]]
            if a2 >= b1 and b2 >= a1:
                uf.union(ids[i], above[j])
            if a2 < b2:
                i += 1
            else:
                j += 1
    boxes: Dict[int, Box] = {}
    for idx, (r, c1, c2) in enumerate(runs):
        root = uf.find(idx)
        b = boxes.get(root)
        boxes[root] = Box(r, c1, r, c2) if b is None else b.union(Box(r, c1, r, c2))
    return list(boxes.values())


def _merge_overlapping(boxes: List[Box]) -> List[Box]:
    """겹치거나 안에 들어간 덩어리를 합친다(합친 결과가 또 겹칠 수 있어 반복)."""
    changed = True
    while changed and len(boxes) > 1:
        changed = False
        boxes.sort(key=lambda b: (b.r1, b.c1))
        out: List[Box] = []
        used = [False] * len(boxes)
        for i, a in enumerate(boxes):
            if used[i]:
                continue
            cur = a
            j = i + 1
            while j < len(boxes) and boxes[j].r1 <= cur.r2:
                if not used[j] and cur.overlaps(boxes[j]):
                    cur = cur.union(boxes[j])
                    used[j] = True
                    changed = True
                j += 1
            out.append(cur)
        boxes = out
    return boxes


def _row_kinds(grid: SheetGrid, r: int, c1: int, c2: int) -> List[Optional[str]]:
    out = []
    for c in range(c1, c2 + 1):
        cell = grid.cells.get((r, c))
        out.append(cell.kind if cell else None)
    return out


def _row_texts(grid: SheetGrid, r: int, c1: int, c2: int) -> List[str]:
    return [(grid.cells[(r, c)].text if (r, c) in grid.cells else "") for c in range(c1, c2 + 1)]


def _last_text_row(grid: SheetGrid, box: Box) -> Optional[int]:
    for r in range(box.r2, box.r1 - 1, -1):
        if any((r, c) in grid.cells for c in range(box.c1, box.c2 + 1)):
            return r
    return None


def _first_text_row(grid: SheetGrid, box: Box, start: Optional[int] = None) -> Optional[int]:
    for r in range(start or box.r1, box.r2 + 1):
        if any((r, c) in grid.cells for c in range(box.c1, box.c2 + 1)):
            return r
    return None


def _same_shape(a: List[Optional[str]], b: List[Optional[str]]) -> bool:
    both = [(x, y) for x, y in zip(a, b) if x and y]
    if not both:
        return False
    same = sum(1 for x, y in both if x == y)
    return same / len(both) >= 0.8


#: 덩어리가 이보다 많으면(값이 흩어진 시트) 빈 줄 이어 붙이기를 하지 않는다. 비교가 제곱으로 는다.
_MAX_BLOCKS_FOR_GAP_MERGE = 2000


def _merge_gaps(grid: SheetGrid, blocks: List[SheetBlock]) -> List[SheetBlock]:
    if len(blocks) > _MAX_BLOCKS_FOR_GAP_MERGE:
        return blocks
    changed = True
    while changed and len(blocks) > 1:
        changed = False
        blocks.sort(key=lambda b: (b.box.r1, b.box.c1))
        for i in range(len(blocks)):
            a = blocks[i]
            for j in range(len(blocks)):
                if i == j:
                    continue
                b = blocks[j]
                A, B = a.box, b.box
                # 가로: 빈 열 한 줄, 같은 행 범위
                if B.c1 - A.c2 == 2 and A.r1 == B.r1 and A.r2 == B.r2:
                    a.box = A.union(B)
                    a.skip_rows |= b.skip_rows
                    del blocks[j]
                    changed = True
                    break
                # 세로: 빈 행 한 줄, 아래가 위 열 범위 안, 첫 행이 위 데이터와 같은 모양
                if B.r1 - A.r2 == 2 and B.c1 >= A.c1 and B.c2 <= A.c2 and A.r2 > A.r1:
                    a_first = _first_text_row(grid, A)
                    a_last = _last_text_row(grid, A)
                    b_first = _first_text_row(grid, B)
                    if a_first is None or a_last is None or b_first is None or a_first == a_last:
                        continue
                    b_texts = _row_texts(grid, b_first, A.c1, A.c2)
                    if b_texts == _row_texts(grid, a_first, A.c1, A.c2):
                        # 같은 머리글이 다시 나온 것(인쇄용 반복 머리글)이면 이어 붙이고 그 행은 버린다
                        a.box = A.union(B)
                        a.skip_rows |= b.skip_rows | {b_first}
                        del blocks[j]
                        changed = True
                        break
                    if _same_shape(_row_kinds(grid, a_last, A.c1, A.c2), _row_kinds(grid, b_first, A.c1, A.c2)):
                        a.box = A.union(B)
                        a.skip_rows |= b.skip_rows
                        del blocks[j]
                        changed = True
                        break
            if changed:
                break
    return blocks


def detect_blocks(grid: SheetGrid) -> List[SheetBlock]:
    """시트를 덩어리로 나눈다(위→아래, 왼쪽→오른쪽 순)."""
    occ = _occupied_by_row(grid)
    if not occ:
        return []
    boxes = _merge_overlapping(_components(occ))
    blocks = _merge_gaps(grid, [SheetBlock(b) for b in boxes])
    blocks.sort(key=lambda b: (b.box.r1, b.box.c1))
    return blocks


# ──────────────────────────────────────────────────────────────────────────────
# 그리기
# ──────────────────────────────────────────────────────────────────────────────

class _MergeIndex:
    def __init__(self, grid: SheetGrid):
        self.at: Dict[Tuple[int, int], Tuple[int, int, int, int]] = {}
        for m in grid.merges:
            r1, c1, r2, c2 = m
            if (r2 - r1 + 1) * (c2 - c1 + 1) > _MAX_MERGE_AREA:
                continue
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    self.at[(r, c)] = m


def _md(text: str) -> str:
    return text.replace("|", "\\|")


def _render_block(grid: SheetGrid, block: SheetBlock, merges: _MergeIndex,
                  by_row: Dict[int, List[int]]) -> Tuple[str, str, str]:
    """덩어리 하나 → (종류, 본문, 표 제목). 종류: table | text | ''(빈 덩어리)."""
    box = block.box
    rows: List[int] = []
    colset: Set[int] = set()
    for r in range(box.r1, box.r2 + 1):
        if r in block.skip_rows:
            continue
        cs = [c for c in by_row.get(r, ()) if box.c1 <= c <= box.c2]
        if cs:
            rows.append(r)
            colset.update(cs)
    if not rows:
        return "", "", ""
    cols = sorted(colset)

    def text_at(r: int, c: int) -> str:
        cell = grid.cells.get((r, c))
        return cell.text if cell else ""

    def kind_at(r: int, c: int) -> Optional[str]:
        cell = grid.cells.get((r, c))
        return cell.kind if cell else None

    # 한 행·한 열 → 글. 파이프로 잇지 않는다(파이프 표 파서가 표로 오인한다).
    if len(rows) == 1 or len(cols) == 1:
        lines = []
        for r in rows:
            vals = [text_at(r, c) for c in cols if text_at(r, c)]
            if vals:
                lines.append(TEXT_JOINER.join(vals))
        return "text", "\n".join(lines), ""

    # 양식: 두 열이 빈 열을 사이에 둔다 → 라벨: 값
    if len(cols) == 2 and cols[1] - cols[0] >= 2 and all(kind_at(r, cols[0]) in (None, "text") for r in rows):
        lines = []
        for r in rows:
            k, v = text_at(r, cols[0]), text_at(r, cols[1])
            if k and v:
                lines.append(f"{k.rstrip(':：').strip()}: {v}")
            elif k or v:
                lines.append(k or v)
        return "text", "\n".join(lines), ""

    def merge_of(r: int, c: int) -> Optional[Tuple[int, int, int, int]]:
        return merges.at.get((r, c))

    def resolved(r: int, c: int) -> str:
        """머리글용: 병합 영역 안이면 병합 값."""
        t = text_at(r, c)
        if t:
            return t
        m = merge_of(r, c)
        return text_at(m[0], m[1]) if m else ""

    # 제목 행: 값이 한 칸뿐인 위쪽 행(표 너비로 병합됐거나, 다음 행이 여러 칸을 가진 경우)
    captions: List[str] = []
    k = 0
    while k < len(rows) - 1 and len(captions) < _MAX_CAPTION_ROWS:
        r = rows[k]
        present = [c for c in cols if text_at(r, c)]
        if len(present) != 1:
            break
        c0 = present[0]
        m = merge_of(r, c0)
        wide_merge = bool(m and (m[3] - m[1] + 1) >= max(2, math.ceil(len(cols) * 0.5)))
        nxt = rows[k + 1]
        next_filled = sum(1 for c in cols if resolved(nxt, c))
        if wide_merge or (c0 == cols[0] and next_filled >= 2):
            captions.append(text_at(r, c0))
            k += 1
            continue
        break

    # 머리글 깊이
    def header_like(r: int) -> bool:
        filled = [c for c in cols if text_at(r, c)]
        if len(filled) < max(2, math.ceil(len(cols) / 2)):
            return False
        return all(kind_at(r, c) == "text" for c in filled)

    H = 0
    if k < len(rows):
        h0 = rows[k]
        starts = [m for m in {merge_of(h0, c) for c in cols} if m and m[0] == h0]
        rowspans = [m[2] - m[0] + 1 for m in starts]
        colspans = [m[3] - m[1] + 1 for m in starts]
        remaining = len(rows) - k
        if rowspans and max(rowspans) > 1:
            last = max(m[2] for m in starts)
            H = sum(1 for r in rows[k:] if r <= last)
        elif colspans and max(colspans) > 1 and remaining > 2 and header_like(rows[k + 1]):
            H = 2
        elif header_like(h0) and remaining > 1:
            H = 1
        H = min(H, _MAX_HEADER_ROWS, max(0, remaining - 1))

    header_rows = rows[k:k + H]
    data_rows = rows[k + H:]
    if header_rows:
        names = []
        for c in cols:
            parts: List[str] = []
            for r in header_rows:
                t = resolved(r, c)
                if t and (not parts or parts[-1] != t):
                    parts.append(t)
            names.append(" ".join(parts) or column_letter(c))
    else:
        names = [column_letter(c) for c in cols]

    lines = ["| " + " | ".join(_md(n) for n in names) + " |",
             "| " + " | ".join("---" for _ in names) + " |"]
    body = 0
    for r in data_rows:
        vals = []
        for c in cols:
            t = text_at(r, c)
            if not t:
                m = merge_of(r, c)
                if m and c == m[1] and r > m[0]:  # 세로 병합 → 아래 행에도 값을 채운다
                    t = text_at(m[0], m[1])
            vals.append(_md(t))
        if any(vals):
            lines.append("| " + " | ".join(vals) + " |")
            body += 1
    caption = " / ".join(captions)
    if body == 0:
        # 데이터 행이 없다 → 표가 아니라 글
        text_lines = list(captions) + [TEXT_JOINER.join(n for n in names if n)]
        return "text", "\n".join(t for t in text_lines if t), ""
    return "table", "\n".join(lines), caption


def render_sheet(grid: SheetGrid) -> List[str]:
    """시트 → 출력 조각 목록(덩어리마다 하나). 표 조각은 ``[Table N] 제목`` 줄을 앞에 단다
    (표가 둘 이상이거나 제목이 있을 때)."""
    blocks = detect_blocks(grid)
    if not blocks:
        return []
    by_row: Dict[int, List[int]] = {}
    for r, c in grid.cells:
        by_row.setdefault(r, []).append(c)
    for cs in by_row.values():
        cs.sort()
    merges = _MergeIndex(grid)
    rendered = [_render_block(grid, b, merges, by_row) for b in blocks]
    rendered = [x for x in rendered if x[0]]
    n_tables = sum(1 for kind, _, _ in rendered if kind == "table")
    parts: List[str] = []
    t_idx = 0
    for kind, content, caption in rendered:
        if kind == "table":
            t_idx += 1
            if n_tables > 1 or caption:
                marker = f"[Table {t_idx}]" + (f" {caption}" if caption else "")
                parts.append(f"{marker}\n{content}")
            else:
                parts.append(content)
        elif content.strip():
            parts.append(content)
    return parts
