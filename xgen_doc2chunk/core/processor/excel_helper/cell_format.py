"""
엑셀 셀 값을 **사용자가 엑셀에서 보는 모양**으로 바꾼다.

셀에는 원시 값(0.153, 1234567, 날짜 일련번호)과 표시 서식(``0.0%``, ``#,##0"원"``,
``yyyy-mm-dd``)이 따로 저장된다. 원시 값을 그대로 쓰면 검색어("15.3%", "1,234,567원")와
맞지 않고 답변에서 단위가 사라지므로, 표시 서식을 적용한 글자를 만든다.

엑셀 서식 문법 전체가 아니라 실무 파일에 나오는 부분을 다룬다:
- 구간(양수;음수;0;텍스트), 색·조건 표시([Red], [>100]) 무시
- 숫자 자리표시(0 # ?), 천 단위 구분(,), 소수 자리, 백분율(%), 지수(E+), 천 단위 나눔(끝의 ,)
- 따옴표 글자("원"), 역슬래시 글자(\\-), 통화 기호([$₩-412]), 자리 채움(_x, *x)
- 날짜·시간(yyyy yy mmmm mmm mm m dddd ddd dd d aaaa aaa hh h mm ss AM/PM)
해석할 수 없는 서식은 일반(General) 표시로 되돌린다.
"""

from __future__ import annotations

import datetime as _dt
import re
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from functools import lru_cache
from typing import List, Optional, Tuple

__all__ = ["format_cell_value", "format_general_number", "is_date_format"]

_BRACKET_RE = re.compile(r"\[[^\]]*\]")

_WEEKDAYS_EN = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
_WEEKDAYS_KO = ["월요일", "화요일", "수요일", "목요일", "금요일", "토요일", "일요일"]
_MONTHS_EN = ["January", "February", "March", "April", "May", "June", "July",
              "August", "September", "October", "November", "December"]


def format_general_number(value: float) -> str:
    """일반(General) 서식: 정수면 소수점 없이, 실수는 유효숫자 15자리(부동소수 잡음 제거)."""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, int):
        return str(value)
    if value != value:  # NaN
        return ""
    if value.is_integer() and abs(value) < 1e15:
        return str(int(value))
    text = f"{value:.15g}"
    if "e" in text:
        mant, exp = text.split("e")
        return f"{mant}E{int(exp):+03d}"
    return text


def _split_sections(fmt: str) -> List[str]:
    """``;`` 로 구간을 나눈다. 따옴표·역슬래시 안의 ``;`` 는 글자다."""
    sections, cur, i, quoted = [], [], 0, False
    while i < len(fmt):
        ch = fmt[i]
        if ch == '"':
            quoted = not quoted
            cur.append(ch)
        elif ch == "\\" and not quoted and i + 1 < len(fmt):
            cur.append(fmt[i:i + 2])
            i += 1
        elif ch == ";" and not quoted:
            sections.append("".join(cur))
            cur = []
        else:
            cur.append(ch)
        i += 1
    sections.append("".join(cur))
    return sections


def _strip_literals(fmt: str) -> str:
    """따옴표·역슬래시 글자와 대괄호를 지운 서식(토큰 판정용)."""
    fmt = re.sub(r'"[^"]*"', "", fmt)
    fmt = re.sub(r"\\.", "", fmt)
    fmt = re.sub(r"[_*].", "", fmt)
    return _BRACKET_RE.sub("", fmt)


@lru_cache(maxsize=512)
def is_date_format(fmt: Optional[str]) -> bool:
    """날짜·시간 서식인가."""
    if not fmt:
        return False
    first = _split_sections(fmt)[0]
    if re.search(r"\[(?:h+|m+|s+)\]", first, re.I):
        return True
    core = _strip_literals(first).lower()
    if core in ("general", "@", ""):
        return False
    return bool(re.search(r"[ydhs]|am/pm|a/p", core)) or (
        "m" in core and not re.search(r"[0#?]", core)
    )


@lru_cache(maxsize=512)
def _tokenize(section: str) -> Tuple[Tuple[str, str], ...]:
    """서식 구간을 (종류, 글자) 토큰으로. 종류: lit(그대로 출력), other(서식 문자)."""
    tokens: List[Tuple[str, str]] = []
    i = 0
    while i < len(section):
        ch = section[i]
        if ch == '"':
            j = section.find('"', i + 1)
            j = len(section) if j < 0 else j
            tokens.append(("lit", section[i + 1:j]))
            i = j + 1
            continue
        if ch == "\\" and i + 1 < len(section):
            tokens.append(("lit", section[i + 1]))
            i += 2
            continue
        if ch == "_" and i + 1 < len(section):  # 자리 맞춤 공백
            tokens.append(("lit", " "))
            i += 2
            continue
        if ch == "*" and i + 1 < len(section):  # 반복 채움은 글자로는 의미 없다
            i += 2
            continue
        if ch == "[":
            j = section.find("]", i)
            j = len(section) if j < 0 else j
            body = section[i + 1:j]
            m = re.match(r"\$([^-]*)", body)
            if m and m.group(1):
                tokens.append(("lit", m.group(1)))
            i = j + 1
            continue
        tokens.append(("other", ch))
        i += 1
    return tuple(tokens)


def _round_half_up(value: Decimal, places: int) -> Decimal:
    q = Decimal(1).scaleb(-places) if places > 0 else Decimal(1)
    return value.quantize(q, rounding=ROUND_HALF_UP)


def _format_number_section(value: float, section: str) -> Optional[str]:
    tokens = list(_tokenize(section))
    chars = "".join(t for k, t in tokens if k == "other")
    if not re.search(r"[0#?]", chars):
        # 자리표시가 없는 구간(예: "-" 만 있는 0 구간)은 글자만 출력
        return "".join(t for k, t in tokens) if tokens else None

    placeholder = set("0#?,.%Ee+-")
    first = last = None
    for idx, (kind, t) in enumerate(tokens):
        if kind == "other" and t in "0#?":
            first = idx if first is None else first
            last = idx
    j = last + 1
    while j < len(tokens) and tokens[j][0] == "other" and tokens[j][1] in placeholder - {"%"}:
        if tokens[j][1] in "+-" and not (j > 0 and tokens[j - 1][1] in "Ee"):
            break
        last = j
        j += 1
    prefix_tokens = tokens[:first]
    suffix_tokens = tokens[last + 1:]
    body = "".join(t for k, t in tokens[first:last + 1])

    percent = sum(1 for k, t in tokens if k == "other" and t == "%")
    v = Decimal(repr(abs(value))) if not isinstance(value, int) else Decimal(abs(value))
    if percent:
        v *= Decimal(100) ** percent

    sci = re.search(r"[Ee][+-]", body)
    if sci:
        mant = body[:sci.start()]
        decimals = len(mant.split(".", 1)[1]) if "." in mant else 0
        exp_digits = max(2, len(body[sci.end():]))
        text = f"{float(v):.{decimals}E}"
        m, e = text.split("E")
        text = f"{m}E{int(e):+0{exp_digits + 1}d}"
    else:
        int_part, _, dec_part = body.partition(".")
        trailing = len(int_part) - len(int_part.rstrip(","))  # 끝에 붙은 , 는 천 단위 나눔
        if trailing:
            v /= Decimal(1000) ** trailing
            int_part = int_part.rstrip(",")
        grouping = "," in int_part
        min_int = int_part.count("0")
        dec_digits = [c for c in dec_part if c in "0#?"]
        max_dec = len(dec_digits)
        min_dec = sum(1 for c in dec_digits if c in "0?")
        try:
            rounded = _round_half_up(v, max_dec)
        except InvalidOperation:
            return None
        text = f"{rounded:.{max_dec}f}" if max_dec else f"{rounded:.0f}"
        ip, _, dp = text.partition(".")
        if dp:
            dp = dp.rstrip("0")
            dp = dp + "0" * max(0, min_dec - len(dp))
        ip = ip.lstrip("0") or ""
        if len(ip) < min_int:
            ip = "0" * (min_int - len(ip)) + ip
        if grouping and ip:
            ip = f"{int(ip):,}"
        text = ip + ("." + dp if dp else "")
        if not text:
            text = "0"
    pre = "".join(t for k, t in prefix_tokens)
    suf = "".join(t for k, t in suffix_tokens)
    return f"{pre}{text}{suf}"


def _format_number(value: float, fmt: str) -> str:
    sections = _split_sections(fmt)
    sec = sections[0]
    neg_sign = value < 0
    if value < 0 and len(sections) >= 2 and sections[1].strip():
        sec, neg_sign = sections[1], False  # 음수 구간이 부호를 직접 그린다
    elif value == 0 and len(sections) >= 3 and sections[2].strip():
        sec = sections[2]
    core = _strip_literals(sec).strip().lower()
    if core in ("general", "") and '"' not in sec:
        text = format_general_number(abs(value) if neg_sign else value)
        return ("-" + text) if neg_sign and not text.startswith("-") else text
    if core == "@":
        return format_general_number(value)
    out = _format_number_section(value, sec)
    if out is None:
        return format_general_number(value)
    if neg_sign:
        out = "-" + out
    return out


def _format_datetime(value: _dt.datetime, fmt: str) -> str:
    section = _split_sections(fmt)[0]
    tokens = _tokenize(section)
    raw = "".join(t if k == "other" else "\0" for k, t in tokens)
    lits = [t for k, t in tokens if k == "lit"]
    has_ampm = bool(re.search(r"(?i)am/pm|a/p", raw))

    parts: List[Tuple[str, str]] = []  # (kind, text)
    i, lit_i = 0, 0
    while i < len(raw):
        ch = raw[i]
        if ch == "\0":
            parts.append(("lit", lits[lit_i]))
            lit_i += 1
            i += 1
            continue
        if raw[i:i + 5].lower() == "am/pm":
            parts.append(("ampm", raw[i:i + 5]))
            i += 5
            continue
        if raw[i:i + 3].lower() == "a/p":
            parts.append(("ap", raw[i:i + 3]))
            i += 3
            continue
        c = ch.lower()
        if c in "ymdhsae":
            j = i
            while j < len(raw) and raw[j].lower() == c:
                j += 1
            parts.append(("tok", raw[i:j].lower()))
            i = j
            continue
        if c == "." and i + 1 < len(raw) and raw[i + 1] == "0":  # 초의 소수(ss.0)
            j = i + 1
            while j < len(raw) and raw[j] == "0":
                j += 1
            parts.append(("frac", raw[i:j]))
            i = j
            continue
        parts.append(("lit", ch))
        i += 1

    out: List[str] = []
    for idx, (kind, t) in enumerate(parts):
        if kind == "lit":
            out.append(t)
            continue
        if kind == "frac":
            frac = value.microsecond / 1_000_000
            out.append(f"{frac:.{len(t) - 1}f}"[1:])
            continue
        if kind in ("ampm", "ap"):
            pm = value.hour >= 12
            out.append(("PM" if pm else "AM") if kind == "ampm" else ("P" if pm else "A"))
            continue
        c, n = t[0], len(t)
        if c == "m":
            # m 이 분인지 월인지: 앞의 h 뒤 또는 뒤의 s 앞이면 분
            prev_tok = next((pt for pk, pt in reversed(parts[:idx]) if pk == "tok"), "")
            next_tok = next((nt for nk, nt in parts[idx + 1:] if nk == "tok"), "")
            if prev_tok.startswith("h") or next_tok.startswith("s"):
                out.append(f"{value.minute:0{min(n, 2)}d}")
            elif n == 1:
                out.append(str(value.month))
            elif n == 2:
                out.append(f"{value.month:02d}")
            elif n == 3:
                out.append(_MONTHS_EN[value.month - 1][:3])
            else:
                out.append(_MONTHS_EN[value.month - 1])
        elif c in "ye":
            out.append(f"{value.year % 100:02d}" if n <= 2 and c == "y" else str(value.year))
        elif c == "d":
            if n == 1:
                out.append(str(value.day))
            elif n == 2:
                out.append(f"{value.day:02d}")
            elif n == 3:
                out.append(_WEEKDAYS_EN[value.weekday()][:3])
            else:
                out.append(_WEEKDAYS_EN[value.weekday()])
        elif c == "a":  # aaa/aaaa: 한국어 요일
            out.append(_WEEKDAYS_KO[value.weekday()][:1] if n == 3 else _WEEKDAYS_KO[value.weekday()])
        elif c == "h":
            hour = (value.hour % 12 or 12) if has_ampm else value.hour
            out.append(f"{hour:0{min(n, 2)}d}")
        elif c == "s":
            out.append(f"{value.second:0{min(n, 2)}d}")
        else:
            out.append(t)
    return "".join(out).strip()


def _excel_serial_to_datetime(serial: float, date1904: bool = False) -> Optional[_dt.datetime]:
    try:
        base = _dt.datetime(1904, 1, 1) if date1904 else _dt.datetime(1899, 12, 30)
        return base + _dt.timedelta(days=float(serial))
    except (OverflowError, ValueError):
        return None


def format_cell_value(
    value,
    number_format: Optional[str] = None,
    *,
    date1904: bool = False,
) -> str:
    """셀 값 하나를 엑셀 표시 문자열로.

    Args:
        value: openpyxl/xlrd 가 준 값 (str, int, float, bool, datetime, date, time, timedelta, None)
        number_format: 셀 표시 서식 (없으면 General)
        date1904: 1904 날짜 체계 (숫자 일련번호를 날짜로 바꿀 때만 쓴다)
    """
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, str):
        return value
    fmt = (number_format or "General").strip() or "General"

    if isinstance(value, _dt.timedelta):
        total = int(value.total_seconds())
        sign = "-" if total < 0 else ""
        total = abs(total)
        return f"{sign}{total // 3600}:{total % 3600 // 60:02d}:{total % 60:02d}"
    if isinstance(value, _dt.time):
        dtv = _dt.datetime.combine(_dt.date(1900, 1, 1), value)
        if is_date_format(fmt):
            try:
                return _format_datetime(dtv, fmt)
            except Exception:
                pass
        return value.strftime("%H:%M:%S" if value.second else "%H:%M")
    if isinstance(value, (_dt.datetime, _dt.date)):
        dtv = value if isinstance(value, _dt.datetime) else _dt.datetime.combine(value, _dt.time())
        if is_date_format(fmt):
            try:
                text = _format_datetime(dtv, fmt)
                if text:
                    return text
            except Exception:
                pass
        if dtv.time() == _dt.time():
            return dtv.strftime("%Y-%m-%d")
        return dtv.strftime("%Y-%m-%d %H:%M:%S" if dtv.second else "%Y-%m-%d %H:%M")
    if isinstance(value, (int, float)):
        if is_date_format(fmt):
            dtv = _excel_serial_to_datetime(value, date1904)
            if dtv is not None:
                return format_cell_value(dtv, fmt)
        try:
            return _format_number(value, fmt)
        except Exception:
            return format_general_number(value)
    return str(value)
