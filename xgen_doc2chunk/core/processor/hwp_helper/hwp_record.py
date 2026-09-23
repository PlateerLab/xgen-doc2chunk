"""
HWP Record 파싱 클래스

본 제품은 한글과컴퓨터의 글 문서 파일(.hwp) 공개 문서를 참고하여 개발하였습니다.
"""
import struct
import logging
from itertools import islice
from typing import Optional

from xgen_doc2chunk.core.processor.hwp_helper.hwp_constants import HWPTAG_PARA_TEXT

logger = logging.getLogger("document-processor")

# HWP 5.0 spec (revision 1.3) table 6 - control characters occupying 8 words.
_INLINE_CONTROLS = frozenset({4, 5, 6, 7, 8, 9, 19, 20})
_EXTENDED_CONTROLS = frozenset({1, 2, 3, 11, 12, 14, 15, 16, 17, 18, 21, 22, 23})


class HwpRecord:
    def __init__(self, tag_id: int, payload: bytes, parent: 'HwpRecord' = None):
        self.tag_id = tag_id
        self.payload = payload
        self.parent = parent
        self.children = []

    def get_next_siblings(self, count=None):
        if not self.parent:
            return []
        try:
            start_idx = self.parent.children.index(self) + 1
            if count is None:
                end_idx = None
            else:
                end_idx = start_idx + count
            return islice(self.parent.children, start_idx, end_idx)
        except ValueError:
            return []

    def get_text(self) -> str:
        """
        Extract text from HWPTAG_PARA_TEXT payload, handling control characters.
        Returns text with \\x0b markers for extended controls.

        Control characters follow the HWP 5.0 spec (revision 1.3, table 6):
        char type (1 word) - 10 line break, 13 paragraph end, 24 hyphen,
        30 non-breaking space, 31 fixed-width space; inline type (8 words) -
        4-9, 19, 20 (9 is TAB); extended type (8 words) - 1-3, 11, 12,
        14-18, 21-23. Ordinary characters are decoded as UTF-16LE runs so a
        surrogate pair stays one character.
        """
        if self.tag_id != HWPTAG_PARA_TEXT:
            return ""

        payload = self.payload
        end = len(payload) - (len(payload) & 1)
        parts = []
        run_start = None
        cursor = 0

        while cursor < end:
            code = payload[cursor] | (payload[cursor + 1] << 8)
            if code >= 32:
                if run_start is None:
                    run_start = cursor
                cursor += 2
                continue
            if run_start is not None:
                parts.append(payload[run_start:cursor].decode('utf-16le', errors='replace'))
                run_start = None

            if code in (10, 13):  # line break / paragraph end
                parts.append('\n')
                cursor += 2
            elif code == 24:  # hyphen
                parts.append('-')
                cursor += 2
            elif code in (30, 31):  # non-breaking / fixed-width space
                parts.append(' ')
                cursor += 2
            elif code == 9:  # TAB is an inline control: 8 words, not 1
                parts.append('\t')
                cursor += 16
            elif code in _INLINE_CONTROLS:
                cursor += 16
            elif code in _EXTENDED_CONTROLS:
                # Code 11 is the standard "Extended Control" marker (for Tables, GSO, etc.)
                if code == 11:
                    parts.append('\x0b')
                cursor += 16
            else:  # remaining char-type controls (0, 25-29 reserved)
                cursor += 2

        if run_start is not None:
            parts.append(payload[run_start:end].decode('utf-16le', errors='replace'))
        return ''.join(parts)

    @staticmethod
    def build_tree(data: bytes) -> 'HwpRecord':
        root = HwpRecord(0, b'')
        pos = 0
        size = len(data)

        # Stack to keep track of parents based on level
        # Level 0 is root children
        # stack[0] = root
        stack = {0: root}

        while pos < size:
            try:
                if pos + 4 > size:
                    break
                header = struct.unpack('<I', data[pos:pos+4])[0]
                pos += 4

                tag_id = header & 0x3FF
                level = (header >> 10) & 0x3FF
                rec_len = (header >> 20) & 0xFFF

                if rec_len == 0xFFF:
                    if pos + 4 > size:
                        break
                    rec_len = struct.unpack('<I', data[pos:pos+4])[0]
                    pos += 4

                if pos + rec_len > size:
                    # Truncated record, stop parsing
                    break

                payload = data[pos:pos+rec_len]
                pos += rec_len

                # Determine parent
                parent = stack.get(level - 1, root)
                if level == 0:
                    parent = root

                # If parent is not in stack (gap in levels), fallback to root or nearest
                if parent is None:
                    # Find nearest lower level
                    for l in range(level - 1, -1, -1):
                        if l in stack:
                            parent = stack[l]
                            break
                    if parent is None:
                        parent = root

                record = HwpRecord(tag_id, payload, parent)
                parent.children.append(record)

                # Update stack for this level
                stack[level] = record

                # Clear deeper levels from stack as we moved to a new node at this level
                keys_to_remove = [k for k in stack.keys() if k > level]
                for k in keys_to_remove:
                    del stack[k]
            except Exception as e:
                logger.debug(f"Error parsing HWP record at pos {pos}: {e}")
                break

        return root
