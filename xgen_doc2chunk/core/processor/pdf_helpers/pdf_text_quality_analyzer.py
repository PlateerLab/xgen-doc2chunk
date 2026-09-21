"""
Text Quality Analyzer for PDF Handler

Analyzes the quality of text extracted from PDF and detects broken text
(encoding issues, missing ToUnicode CMap, etc.) to determine whether OCR fallback is needed.

=============================================================================
Characteristics of Broken Text:
=============================================================================
1. Contains many Private Use Area (PUA) characters: U+E000 ~ U+F8FF
2. Replacement Character: U+FFFD (�)
3. Invalid Korean character combinations (only consonants/vowels in sequence)
4. Meaningless Korean syllable sequences (random combinations, not real words)
5. Mixture of CJK characters with PUA/control characters
6. CJK Compatibility characters used instead of normal punctuation
7. Fragmented text where each character is on a separate line

=============================================================================
Resolution Strategy:
=============================================================================
1. Calculate text quality score (0.0 ~ 1.0)
2. For fragmented text: Reconstruct using character position data
3. For CJK Compatibility characters: Map to correct characters
4. Perform OCR fallback only if reconstruction fails
"""

import logging
import re
import unicodedata
from typing import List, Dict, Tuple, Optional, Set
from dataclasses import dataclass, field

import xgen_pdf as fitz
from PIL import Image
import pytesseract

logger = logging.getLogger(__name__)


# ============================================================================
# CJK Compatibility Character Mapping
# ============================================================================

# Map CJK Compatibility characters to their intended characters
# These occur when Word documents are converted to PDF with font issues
CJK_COMPAT_CHAR_MAP = {
    # Parentheses
    '\u33D9': '(',      # ㏙ → (
    '\u33DA': ')',      # ㏚ → )
    
    # Brackets (section markers)
    '\u33DB': '[',      # ㏛ → [ (or could be 【)
    '\u33DC': ']',      # ㏜ → ] (or could be 】)
    '\u33DD': '[',      # ㏝ → [ (section start)
    '\u33DE': ']',      # ㏞ → ] (section end)
    
    # Arrows and connectors
    '\u3711': '→',      # 㜑 → arrow
    '\u36A8': '/',      # 㚨 → / or +
    '\u36F3': '→',      # 㛳 → arrow (Word→PDF conversion often maps arrow to this)
    '\u3689': '+',      # 㚉 → + (plus sign, e.g., Vector + Graph)
    
    # Range indicator
    '\u33CA': '~',      # ㏊ → ~ (range, e.g., 2~6개월)
    
    # Quotation marks
    '\u3431': '"',      # 㐱 → opening quote
    '\u3432': '"',      # 㐲 → closing quote
    '\u3433': '"',      # 㐳 → opening quote
    '\u3434': '"',      # 㐴 → closing quote
    '\u3443': '"',      # 㑃 → quote
}


class TextQualityConfig:
    """Text quality analysis configuration."""
    
    # Quality threshold
    QUALITY_THRESHOLD = 0.7           # OCR fallback if below this value (raised from 0.5 to 0.7)
    MIN_TEXT_LENGTH = 10              # Minimum text length for quality analysis
    
    # PUA-based threshold (force OCR if PUA ratio is above this)
    PUA_RATIO_THRESHOLD = 0.1         # 10% or more triggers OCR
    
    # PUA (Private Use Area) ranges
    PUA_RANGES = [
        (0xE000, 0xF8FF),     # BMP Private Use Area
        (0xF0000, 0xFFFFD),   # Supplementary PUA-A
        (0x100000, 0x10FFFD), # Supplementary PUA-B
    ]
    
    # CJK Compatibility ranges (often indicates broken text from Word->PDF conversion)
    # These are unit symbols that are rarely used in normal text but appear when
    # character encoding is broken (e.g., parentheses becoming ㏙, ㏚, etc.)
    CJK_COMPAT_RANGES = [
        (0x3300, 0x33FF),     # CJK Compatibility (squared Katakana, units)
        (0x3200, 0x32FF),     # Enclosed CJK Letters and Months
        (0x3700, 0x37FF),     # CJK Extension A (rarely used Hanja)
    ]
    
    # Control characters and special characters
    CONTROL_RANGES = [
        (0x0000, 0x001F),     # C0 controls
        (0x007F, 0x009F),     # C1 controls
        (0xFFF0, 0xFFFF),     # Specials
    ]
    
    # OCR settings
    OCR_LANG = 'kor+eng'
    OCR_DPI = 300
    OCR_SCALE = 3.0
    
    # Korean syllable ranges
    HANGUL_SYLLABLE_RANGE = (0xAC00, 0xD7A3)
    HANGUL_JAMO_RANGE = (0x1100, 0x11FF)
    HANGUL_COMPAT_JAMO_RANGE = (0x3130, 0x318F)
    
    # Quality analysis weights
    WEIGHT_PUA = 0.4              # PUA character ratio weight
    WEIGHT_REPLACEMENT = 0.3      # Replacement character weight
    WEIGHT_VALID_RATIO = 0.3      # Valid character ratio weight
    WEIGHT_CJK_COMPAT = 0.5       # CJK Compatibility character weight (broken text indicator)
    
    # Fragmented text detection settings
    # When each line has only 1-2 characters, it indicates conversion issue
    FRAGMENTED_TEXT_THRESHOLD = 0.5  # If >50% of lines have <=2 chars, text is fragmented
    FRAGMENTED_LINE_CHAR_LIMIT = 3   # Lines with <= this many chars are considered fragmented
    MIN_LINES_FOR_FRAGMENTED_CHECK = 5  # Minimum lines needed to check for fragmentation


# ============================================================================
# Data Classes
# ============================================================================

@dataclass
class TextQualityResult:
    """Text quality analysis result."""
    quality_score: float          # 0.0 ~ 1.0 (higher is better quality)
    total_chars: int              # Total character count
    pua_count: int                # PUA character count
    replacement_count: int        # Replacement character count
    valid_chars: int              # Valid character count (Korean, English, digits)
    control_chars: int            # Control character count
    cjk_compat_count: int = 0     # CJK Compatibility character count (broken text indicator)
    is_fragmented: bool = False   # Whether text is fragmented (char-by-char line breaks)
    needs_ocr: bool = False       # Whether OCR is needed
    details: Dict = None          # Detailed information
    
    def __post_init__(self):
        if self.details is None:
            self.details = {}


@dataclass
class PageTextAnalysis:
    """Page text analysis result."""
    page_num: int
    quality_result: TextQualityResult
    text_blocks: List[Dict]       # Individual text block information
    problem_regions: List[Tuple[float, float, float, float]]  # Bounding boxes of problematic regions
    ocr_text: Optional[str] = None  # OCR result (if performed)


# ============================================================================
# Text Quality Analyzer
# ============================================================================

class TextQualityAnalyzer:
    """
    Text Quality Analyzer.
    
    Analyzes the quality of text extracted from PDF and
    detects broken text to determine whether OCR fallback is needed.
    """
    
    def __init__(self, page, page_num: int):
        """
        Args:
            page: PyMuPDF page object
            page_num: Page number (0-indexed)
        """
        self.page = page
        self.page_num = page_num
        self.page_width = page.rect.width
        self.page_height = page.rect.height
        
    def analyze_page(self) -> PageTextAnalysis:
        """
        Analyze text quality for the entire page.
        
        Returns:
            PageTextAnalysis object
        """
        # Extract text dictionary
        text_dict = self.page.get_text("dict", sort=True)
        blocks = text_dict.get("blocks", [])
        
        all_text = []
        text_blocks = []
        problem_regions = []
        
        # Count lines to detect fragmented text pattern
        # conversion issue where each char is a separate line)
        total_lines = 0
        total_chars = 0
        
        for block in blocks:
            if block.get("type") != 0:  # Text blocks only
                continue
            
            block_bbox = block.get("bbox", (0, 0, 0, 0))
            block_text = []
            
            for line in block.get("lines", []):
                total_lines += 1
                for span in line.get("spans", []):
                    text = span.get("text", "")
                    if text:
                        total_chars += len(text.strip())
                        block_text.append(text)
                        all_text.append(text)
            
            if block_text:
                combined_text = " ".join(block_text)
                quality = self.analyze_text(combined_text)
                
                text_blocks.append({
                    'bbox': block_bbox,
                    'text': combined_text,
                    'quality': quality
                })
                
                # Record low quality regions
                if quality.needs_ocr:
                    problem_regions.append(block_bbox)
        
        # Analyze overall text quality
        full_text = " ".join(all_text)
        overall_quality = self.analyze_text(full_text)
        
        # Detect fragmented text at page level
        # If average chars per line is very low, text is likely fragmented
        if total_lines > 0 and total_chars > 0:
            avg_chars_per_line = total_chars / total_lines
            # If average is less than 15 chars per line, text is fragmented
            page_is_fragmented = avg_chars_per_line < 15 and total_lines >= TextQualityConfig.MIN_LINES_FOR_FRAGMENTED_CHECK
            
            if page_is_fragmented:
                logger.info(
                    f"[QualityAnalyzer] Page {self.page_num + 1}: "
                    f"Detected fragmented text (avg {avg_chars_per_line:.1f} chars/line, {total_lines} lines)"
                )
                # Update overall quality to reflect fragmented status
                overall_quality = TextQualityResult(
                    quality_score=max(0.0, overall_quality.quality_score - 0.5),
                    total_chars=overall_quality.total_chars,
                    pua_count=overall_quality.pua_count,
                    replacement_count=overall_quality.replacement_count,
                    valid_chars=overall_quality.valid_chars,
                    control_chars=overall_quality.control_chars,
                    cjk_compat_count=overall_quality.cjk_compat_count,
                    is_fragmented=True,  # Mark as fragmented
                    needs_ocr=True,  # Trigger reconstruction
                    details={
                        **overall_quality.details,
                        'is_fragmented': True,
                        'avg_chars_per_line': avg_chars_per_line,
                        'total_lines': total_lines,
                    }
                )
        
        return PageTextAnalysis(
            page_num=self.page_num,
            quality_result=overall_quality,
            text_blocks=text_blocks,
            problem_regions=problem_regions
        )
    
    def analyze_text(self, text: str) -> TextQualityResult:
        """
        Analyze text quality.
        
        Args:
            text: Text to analyze
            
        Returns:
            TextQualityResult object
        """
        if not text or len(text) < TextQualityConfig.MIN_TEXT_LENGTH:
            return TextQualityResult(
                quality_score=1.0,  # Treat as OK if text is empty or too short
                total_chars=len(text),
                pua_count=0,
                replacement_count=0,
                valid_chars=len(text),
                control_chars=0,
                cjk_compat_count=0,
                is_fragmented=False,
                needs_ocr=False,
                details={'reason': 'text_too_short'}
            )
        
        total_chars = len(text)
        pua_count = 0
        replacement_count = 0
        control_count = 0
        cjk_compat_count = 0  # CJK Compatibility character count
        valid_chars = 0  # Korean, English, digits, spaces, basic punctuation
        
        # Character-by-character analysis
        for char in text:
            code = ord(char)
            
            # PUA check
            if self._is_pua(code):
                pua_count += 1
                continue
            
            # CJK Compatibility check (broken text indicator)
            if self._is_cjk_compat(code):
                cjk_compat_count += 1
                continue
            
            # Replacement character check
            if code == 0xFFFD:
                replacement_count += 1
                continue
            
            # Control character check
            if self._is_control(code):
                control_count += 1
                continue
            
            # Valid character check
            if self._is_valid_char(char, code):
                valid_chars += 1
        
        # Check for fragmented text pattern (char-by-char line breaks)
        is_fragmented = self._is_fragmented_text(text)
        
        # Calculate quality score
        quality_score = self._calculate_quality_score(
            total_chars=total_chars,
            pua_count=pua_count,
            replacement_count=replacement_count,
            valid_chars=valid_chars,
            cjk_compat_count=cjk_compat_count,
            is_fragmented=is_fragmented
        )
        
        # Determine OCR necessity
        pua_ratio = pua_count / total_chars if total_chars > 0 else 0
        cjk_compat_ratio = cjk_compat_count / total_chars if total_chars > 0 else 0
        needs_ocr = (
            quality_score < TextQualityConfig.QUALITY_THRESHOLD or
            pua_ratio >= TextQualityConfig.PUA_RATIO_THRESHOLD or
            cjk_compat_ratio >= 0.05 or  # 5% or more CJK compat chars triggers OCR
            is_fragmented  # Fragmented text always needs OCR
        )
        
        return TextQualityResult(
            quality_score=quality_score,
            total_chars=total_chars,
            pua_count=pua_count,
            replacement_count=replacement_count,
            valid_chars=valid_chars,
            control_chars=control_count,
            cjk_compat_count=cjk_compat_count,
            is_fragmented=is_fragmented,
            needs_ocr=needs_ocr,
            details={
                'pua_ratio': pua_count / total_chars if total_chars > 0 else 0,
                'replacement_ratio': replacement_count / total_chars if total_chars > 0 else 0,
                'valid_ratio': valid_chars / total_chars if total_chars > 0 else 0,
                'cjk_compat_ratio': cjk_compat_count / total_chars if total_chars > 0 else 0,
                'is_fragmented': is_fragmented,
            }
        )
    
    def _is_pua(self, code: int) -> bool:
        """Check if character is in Private Use Area."""
        for start, end in TextQualityConfig.PUA_RANGES:
            if start <= code <= end:
                return True
        return False
    
    def _is_cjk_compat(self, code: int) -> bool:
        """
        Check if character is in CJK Compatibility range.
        
        These characters often indicate broken text from Word->PDF conversion
        where parentheses, brackets, and other symbols are incorrectly mapped
        to CJK Compatibility characters (e.g., U+3319 for '(', U+331A for ')').
        """
        for start, end in TextQualityConfig.CJK_COMPAT_RANGES:
            if start <= code <= end:
                return True
        return False
    
    def _is_fragmented_text(self, text: str) -> bool:
        """
        Detect fragmented text pattern where each line has only 1-2 characters.
        
        This pattern occurs when Word documents with special layouts
        (text boxes, vertical text, etc.) are converted to PDF,
        resulting in characters being stored as separate lines.
        
        Example of fragmented text:
            '현\n재\n시\n장\n에\n대\n한\n이\n해'
        Should be: '현재 시장에 대한 이해'
        
        Args:
            text: Text to analyze
            
        Returns:
            True if text appears to be fragmented
        """
        lines = text.split('\n')
        
        # Need minimum number of lines to detect pattern
        if len(lines) < TextQualityConfig.MIN_LINES_FOR_FRAGMENTED_CHECK:
            return False
        
        # Count lines with few characters (excluding empty lines)
        non_empty_lines = [line for line in lines if line.strip()]
        if not non_empty_lines:
            return False
        
        short_line_count = sum(
            1 for line in non_empty_lines 
            if len(line.strip()) <= TextQualityConfig.FRAGMENTED_LINE_CHAR_LIMIT
        )
        
        fragmented_ratio = short_line_count / len(non_empty_lines)
        
        return fragmented_ratio >= TextQualityConfig.FRAGMENTED_TEXT_THRESHOLD
    
    def _is_control(self, code: int) -> bool:
        """Check if character is a control character."""
        for start, end in TextQualityConfig.CONTROL_RANGES:
            if start <= code <= end:
                return True
        return False
    
    def _is_valid_char(self, char: str, code: int) -> bool:
        """Check if character is valid (Korean, English, digits, spaces, basic punctuation)."""
        # Whitespace
        if char.isspace():
            return True
        
        # ASCII alphanumeric
        if char.isalnum() and code < 128:
            return True
        
        # Korean syllables
        if TextQualityConfig.HANGUL_SYLLABLE_RANGE[0] <= code <= TextQualityConfig.HANGUL_SYLLABLE_RANGE[1]:
            return True
        
        # Korean Jamo
        if TextQualityConfig.HANGUL_JAMO_RANGE[0] <= code <= TextQualityConfig.HANGUL_JAMO_RANGE[1]:
            return True
        
        # Korean compatibility Jamo
        if TextQualityConfig.HANGUL_COMPAT_JAMO_RANGE[0] <= code <= TextQualityConfig.HANGUL_COMPAT_JAMO_RANGE[1]:
            return True
        
        # Basic punctuation
        if char in '.,!?;:\'"()[]{}-–—…·•':
            return True
        
        # CJK characters (Chinese, Japanese)
        if 0x4E00 <= code <= 0x9FFF:  # CJK Unified Ideographs
            return True
        
        # Japanese Hiragana/Katakana
        if 0x3040 <= code <= 0x30FF:
            return True
        
        return False
    
    def _calculate_quality_score(
        self,
        total_chars: int,
        pua_count: int,
        replacement_count: int,
        valid_chars: int,
        cjk_compat_count: int = 0,
        is_fragmented: bool = False
    ) -> float:
        """Calculate quality score (0.0 ~ 1.0)."""
        if total_chars == 0:
            return 1.0
        
        # Calculate ratios
        pua_ratio = pua_count / total_chars
        replacement_ratio = replacement_count / total_chars
        valid_ratio = valid_chars / total_chars
        cjk_compat_ratio = cjk_compat_count / total_chars
        
        # Calculate weighted score
        # Score decreases with more PUA chars, more replacement chars, lower valid ratio
        score = 1.0
        
        # PUA character penalty (more = lower score)
        score -= pua_ratio * TextQualityConfig.WEIGHT_PUA * 2
        
        # Replacement character penalty
        score -= replacement_ratio * TextQualityConfig.WEIGHT_REPLACEMENT * 3
        
        # CJK Compatibility character penalty (broken text indicator)
        score -= cjk_compat_ratio * TextQualityConfig.WEIGHT_CJK_COMPAT * 3
        
        # Fragmented text penalty (severe quality issue)
        if is_fragmented:
            score -= 0.5  # Major penalty for fragmented text
        
        # Valid character ratio adjustment
        score = score * (0.5 + valid_ratio * 0.5)
        
        return max(0.0, min(1.0, score))


# ============================================================================
# Page OCR Fallback Engine
# ============================================================================

class PageOCRFallbackEngine:
    """
    Page OCR Fallback Engine.
    
    Performs OCR on the entire page or specific regions
    for pages with low text quality.
    """
    
    def __init__(self, page, page_num: int):
        """
        Args:
            page: PyMuPDF page object
            page_num: Page number (0-indexed)
        """
        self.page = page
        self.page_num = page_num
        self.page_width = page.rect.width
        self.page_height = page.rect.height
    
    def ocr_full_page(self) -> str:
        """
        Perform OCR on the entire page.
        
        Returns:
            Text extracted via OCR
        """
        try:
            # Render page at high resolution
            mat = fitz.Matrix(TextQualityConfig.OCR_SCALE, TextQualityConfig.OCR_SCALE)
            pix = self.page.get_pixmap(matrix=mat)
            
            # Convert to PIL Image
            import io
            img_data = pix.tobytes("png")
            img = Image.open(io.BytesIO(img_data))
            
            # Perform OCR (Korean priority)
            ocr_config = '--psm 3 --oem 3'  # Automatic page segmentation + LSTM OCR
            text = pytesseract.image_to_string(
                img,
                lang=TextQualityConfig.OCR_LANG,
                config=ocr_config
            )
            
            # OCR post-processing: noise removal
            text = self._postprocess_ocr_text(text)
            
            logger.info(f"[PageOCR] Page {self.page_num + 1}: OCR extracted {len(text)} chars")
            return text.strip()
            
        except Exception as e:
            logger.error(f"[PageOCR] Page {self.page_num + 1} OCR failed: {e}")
            return ""
    
    def _postprocess_ocr_text(self, text: str) -> str:
        """
        Post-process OCR results.
        
        - Remove lines consisting only of special symbols
        - Remove meaningless short lines
        - Clean up repeated characters
        - Remove OCR noise patterns
        """
        if not text:
            return ""
        
        lines = text.split('\n')
        cleaned_lines = []
        
        # OCR noise patterns (text incorrectly recognized from background graphics)
        noise_patterns = [
            r'^[ri\-—maOANIUTLOG\s]+$',  # Noise from circular background graphics
            r'^[0-9"\'\[\]\(\)°\s]{1,5}$',  # Short number/symbol combinations
            r'^[A-Za-z\-—\s]{3,}$',  # Meaningless English combinations (when no Korean)
            r'^‥+\s*$',  # Only dotted lines
            r'^\s*[°·•○●□■◇◆△▲▽▼]+\s*$',  # Only symbols
        ]
        
        import re
        
        for line in lines:
            line = line.strip()
            
            # Skip empty lines
            if not line:
                continue
            
            # Remove lines consisting only of special symbols
            if all(c in '.,;:!?@#$%^&*()[]{}|\\/<>~`\'"-_+=°·•○●□■◇◆△▲▽▼' or c.isspace() for c in line):
                continue
            
            # Check noise patterns
            is_noise = False
            for pattern in noise_patterns:
                if re.match(pattern, line, re.IGNORECASE):
                    is_noise = True
                    break
            if is_noise:
                continue
            
            # Prioritize keeping lines with Korean
            korean_count = sum(1 for c in line if '가' <= c <= '힣')
            if korean_count > 0:
                cleaned_lines.append(line)
                continue
            
            # For English-only lines, check if meaningful
            alpha_count = sum(1 for c in line if c.isalpha())
            total_len = len(line.replace(' ', ''))
            
            if total_len > 0:
                meaningful_ratio = alpha_count / total_len
                # Keep only if meaningful characters >= 50% and at least 3 characters
                if meaningful_ratio >= 0.5 and alpha_count >= 3:
                    # Keep uppercase abbreviations (PLATEER, IDT, etc.)
                    if line.isupper() or any(word.isupper() and len(word) >= 2 for word in line.split()):
                        cleaned_lines.append(line)
                    # Regular English text (Insight Report, etc.)
                    elif any(c.islower() for c in line):
                        cleaned_lines.append(line)
        
        return '\n'.join(cleaned_lines)
    
    def ocr_region(self, bbox: Tuple[float, float, float, float]) -> str:
        """
        Perform OCR on a specific region.
        
        Args:
            bbox: Region coordinates (x0, y0, x1, y1)
            
        Returns:
            Text extracted via OCR
        """
        try:
            x0, y0, x1, y1 = bbox
            
            # Add padding
            padding = 10
            clip = fitz.Rect(
                max(0, x0 - padding),
                max(0, y0 - padding),
                min(self.page_width, x1 + padding),
                min(self.page_height, y1 + padding)
            )
            
            # Render region at high resolution
            mat = fitz.Matrix(TextQualityConfig.OCR_SCALE, TextQualityConfig.OCR_SCALE)
            pix = self.page.get_pixmap(matrix=mat, clip=clip)
            
            # Convert to PIL Image
            import io
            img_data = pix.tobytes("png")
            img = Image.open(io.BytesIO(img_data))
            
            # Perform OCR
            ocr_config = '--psm 6 --oem 3'  # Uniform text block + LSTM
            text = pytesseract.image_to_string(
                img,
                lang=TextQualityConfig.OCR_LANG,
                config=ocr_config
            )
            
            # OCR post-processing
            text = self._postprocess_ocr_text(text)
            
            return text.strip()
            
        except Exception as e:
            logger.warning(f"[PageOCR] Region OCR failed for {bbox}: {e}")
            return ""
    
    def ocr_problem_regions(
        self,
        problem_regions: List[Tuple[float, float, float, float]]
    ) -> Dict[Tuple, str]:
        """
        Perform OCR on problematic regions.
        
        Args:
            problem_regions: List of bounding boxes for problematic regions
            
        Returns:
            Dictionary mapping {bbox: ocr_text}
        """
        results = {}
        
        for bbox in problem_regions:
            text = self.ocr_region(bbox)
            if text:
                results[bbox] = text
        
        return results


# ============================================================================
# Integrated Text Extractor with Quality Check
# ============================================================================

class QualityAwareTextExtractor:
    """
    Quality-Aware Text Extractor.
    
    Analyzes text quality and performs OCR fallback when necessary
    to always extract high-quality text.
    """
    
    def __init__(self, page, page_num: int, quality_threshold: float = None):
        """
        Args:
            page: PyMuPDF page object
            page_num: Page number (0-indexed)
            quality_threshold: Quality threshold (default: TextQualityConfig.QUALITY_THRESHOLD)
        """
        self.page = page
        self.page_num = page_num
        self.quality_threshold = quality_threshold or TextQualityConfig.QUALITY_THRESHOLD
        
        self.analyzer = TextQualityAnalyzer(page, page_num)
        self.ocr_engine = PageOCRFallbackEngine(page, page_num)
    
    def extract(self) -> Tuple[str, PageTextAnalysis]:
        """
        Extract text with quality consideration.
        
        Returns:
            Tuple of (extracted text, analysis result)
        """
        # 1. Analyze page text quality
        analysis = self.analyzer.analyze_page()
        
        logger.debug(
            f"[QualityAware] Page {self.page_num + 1}: "
            f"quality={analysis.quality_result.quality_score:.2f}, "
            f"pua={analysis.quality_result.pua_count}, "
            f"valid={analysis.quality_result.valid_chars}"
        )
        
        # 2. Return existing text if quality is good
        if not analysis.quality_result.needs_ocr:
            # Extract text using standard method
            text = self.page.get_text("text")
            return text, analysis
        
        # 3. Try text reconstruction first (before OCR)
        # This is more reliable than OCR for fragmented text from Word->PDF conversion
        if analysis.quality_result.is_fragmented or analysis.quality_result.cjk_compat_count > 0:
            logger.info(
                f"[QualityAware] Page {self.page_num + 1}: "
                f"Attempting text reconstruction "
                f"(fragmented={analysis.quality_result.is_fragmented}, "
                f"cjk_compat={analysis.quality_result.cjk_compat_count})"
            )
            
            reconstructor = FragmentedTextReconstructor(self.page, self.page_num)
            reconstructed_text = reconstructor.reconstruct()
            
            if reconstructed_text:
                # Apply CJK Compatibility character mapping
                cleaned_text = apply_cjk_compat_mapping(reconstructed_text)
                analysis.ocr_text = f"[Reconstructed] {len(cleaned_text)} chars"
                return cleaned_text, analysis
        
        # 4. OCR fallback if reconstruction fails
        logger.info(
            f"[QualityAware] Page {self.page_num + 1}: "
            f"Quality too low ({analysis.quality_result.quality_score:.2f}), "
            f"falling back to OCR"
        )
        
        # If few problem regions, OCR only those regions
        if len(analysis.problem_regions) <= 3 and len(analysis.problem_regions) > 0:
            # OCR only problem regions
            ocr_results = self.ocr_engine.ocr_problem_regions(analysis.problem_regions)
            
            # Replace problem region text with OCR results
            text = self._merge_ocr_results(analysis, ocr_results)
            analysis.ocr_text = str(ocr_results)
        else:
            # Full page OCR
            text = self.ocr_engine.ocr_full_page()
            analysis.ocr_text = text
        
        return text, analysis
    
    def _merge_ocr_results(
        self,
        analysis: PageTextAnalysis,
        ocr_results: Dict[Tuple, str]
    ) -> str:
        """
        Merge existing text with OCR results.
        
        Uses existing text for good quality blocks,
        replaces problematic blocks with OCR results.
        """
        merged_parts = []
        
        for block in analysis.text_blocks:
            bbox = tuple(block['bbox'])
            quality = block['quality']
            
            if quality.needs_ocr and bbox in ocr_results:
                # Use OCR result
                merged_parts.append(ocr_results[bbox])
            else:
                # Use existing text
                merged_parts.append(block['text'])
        
        return "\n".join(merged_parts)


# ============================================================================
# Fragmented Text Reconstructor
# ============================================================================

class FragmentedTextReconstructor:
    """
    Reconstructs fragmented text from PDF pages.
    
    When Word documents with special layouts (text boxes, vertical text, etc.)
    are converted to PDF, characters may be stored as separate lines.
    This class reconstructs the text by analyzing character positions.
    
    Example:
        Input: '현\\n재\\n시\\n장\\n에\\n대\\n한\\n이\\n해'
        Output: '현재 시장에 대한 이해'
    """
    
    def __init__(self, page, page_num: int, y_tolerance: float = 3.0, 
                 exclude_bboxes: List[Tuple[float, float, float, float]] = None):
        """
        Args:
            page: PyMuPDF page object
            page_num: Page number (0-indexed)
            y_tolerance: Y coordinate tolerance for same-line detection
            exclude_bboxes: List of bounding boxes to exclude (e.g., table regions)
        """
        self.page = page
        self.page_num = page_num
        self.y_tolerance = y_tolerance
        self.exclude_bboxes = exclude_bboxes or []
    
    def reconstruct(self) -> str:
        """
        Reconstruct fragmented text using character position data.
        
        Returns:
            Reconstructed text with proper line breaks
        """
        try:
            # Extract character-level position data
            raw_dict = self.page.get_text("rawdict")
            all_chars = self._extract_chars(raw_dict)
            
            if not all_chars:
                logger.warning(f"[Reconstruct] Page {self.page_num + 1}: No characters found")
                return ""
            
            # Group characters by Y coordinate (same line)
            lines_by_y = self._group_by_y(all_chars)
            
            # Sort each line by X coordinate and build text
            reconstructed_lines = self._build_lines(lines_by_y)
            
            result = "\n".join(reconstructed_lines)
            
            logger.info(
                f"[Reconstruct] Page {self.page_num + 1}: "
                f"Reconstructed {len(all_chars)} chars into {len(reconstructed_lines)} lines"
            )
            
            return result
            
        except Exception as e:
            logger.error(f"[Reconstruct] Page {self.page_num + 1} failed: {e}")
            return ""
    
    def reconstruct_with_sections(self) -> List[Dict]:
        """
        Reconstruct fragmented text, split into sections by table positions.
        
        This method returns multiple text sections with their Y-coordinate ranges,
        allowing proper positioning relative to tables.
        
        Returns:
            List of dicts: [{'text': str, 'y_start': float, 'y_end': float}, ...]
        """
        try:
            raw_dict = self.page.get_text("rawdict")
            all_chars = self._extract_chars(raw_dict)
            
            if not all_chars:
                logger.warning(f"[Reconstruct] Page {self.page_num + 1}: No characters found")
                return []
            
            # Group characters by Y coordinate
            lines_by_y = self._group_by_y(all_chars)
            
            if not lines_by_y:
                return []
            
            # Get sorted Y positions of tables (exclusion regions)
            table_y_ranges = []
            for bbox in self.exclude_bboxes:
                table_y_ranges.append((bbox[1], bbox[3]))  # (y_start, y_end)
            table_y_ranges.sort(key=lambda x: x[0])
            
            if not table_y_ranges:
                # No tables - return single section
                section_text = self._build_section_text(list(lines_by_y.keys()), lines_by_y)
                if section_text.strip():
                    sorted_ys = sorted(lines_by_y.keys())
                    return [{
                        'text': section_text,
                        'y_start': sorted_ys[0],
                        'y_end': sorted_ys[-1]
                    }]
                return []
            
            # Split lines into sections based on table positions
            # Key insight: when we skip from a Y before table to a Y after table,
            # we need to split the section
            sections = []
            current_section_lines = []
            current_y_start = None
            current_y_end = None
            
            sorted_ys = sorted(lines_by_y.keys())
            
            for y in sorted_ys:
                # Check if we're jumping over a table
                should_split = False
                if current_y_end is not None:
                    for table_y_start, table_y_end in table_y_ranges:
                        # If previous line was before table start AND current line is after table end
                        # (meaning we jumped over the table)
                        if current_y_end < table_y_start and y > table_y_end:
                            should_split = True
                            break
                
                if should_split and current_section_lines:
                    # Save current section (text BEFORE the table)
                    section_text = self._build_section_text(current_section_lines, lines_by_y)
                    if section_text.strip():
                        sections.append({
                            'text': section_text,
                            'y_start': current_y_start,
                            'y_end': current_y_end
                        })
                    current_section_lines = []
                    current_y_start = None
                
                # Add line to current section
                current_section_lines.append(y)
                if current_y_start is None:
                    current_y_start = y
                current_y_end = y
            
            # Don't forget the last section (text AFTER the last table or all text if no split)
            if current_section_lines:
                section_text = self._build_section_text(current_section_lines, lines_by_y)
                if section_text.strip():
                    sections.append({
                        'text': section_text,
                        'y_start': current_y_start,
                        'y_end': current_y_end
                    })
            
            logger.info(
                f"[Reconstruct] Page {self.page_num + 1}: "
                f"Split into {len(sections)} sections around {len(table_y_ranges)} tables"
            )
            
            return sections
            
        except Exception as e:
            logger.error(f"[Reconstruct] Page {self.page_num + 1} sections failed: {e}")
            return []
    
    def _build_section_text(self, y_positions: List[float], lines_by_y: Dict) -> str:
        """Build text from a list of Y positions."""
        lines = []
        for y in sorted(y_positions):
            chars = lines_by_y.get(y, [])
            chars_sorted = sorted(chars, key=lambda c: c['bbox'][0])
            
            if not chars_sorted:
                continue
            
            line_text = ""
            prev_x_end = None
            
            for char_info in chars_sorted:
                x_start = char_info['bbox'][0]
                char = char_info['c']
                
                if prev_x_end is not None:
                    gap = x_start - prev_x_end
                    avg_char_width = char_info['size'] * 0.5
                    if gap > avg_char_width * 0.5:
                        line_text += " "
                
                line_text += char
                prev_x_end = char_info['bbox'][2]
            
            if line_text.strip():
                lines.append(line_text)
        
        return "\n".join(lines)
    
    def _extract_chars(self, raw_dict: Dict) -> List[Dict]:
        """Extract all characters with position info from rawdict.
        
        Characters inside exclude_bboxes (e.g., table regions) are filtered out.
        """
        all_chars = []
        
        for block in raw_dict.get('blocks', []):
            if block.get('type') != 0:  # Text blocks only
                continue
            
            for line in block.get('lines', []):
                for span in line.get('spans', []):
                    font = span.get('font', '')
                    size = span.get('size', 0)
                    
                    for char in span.get('chars', []):
                        char_bbox = char.get('bbox', [0, 0, 0, 0])
                        
                        # Skip characters inside excluded regions (e.g., tables)
                        if self._is_inside_excluded_bbox(char_bbox):
                            continue
                        
                        char_info = {
                            'c': char.get('c', ''),
                            'bbox': char_bbox,
                            'origin': char.get('origin', [0, 0]),
                            'font': font,
                            'size': size,
                        }
                        all_chars.append(char_info)
        
        return all_chars
    
    def _is_inside_excluded_bbox(self, char_bbox: List[float]) -> bool:
        """Check if character is inside any excluded bbox.
        
        Args:
            char_bbox: Character bounding box [x0, y0, x1, y1]
            
        Returns:
            True if character center is inside any excluded region
        """
        if not self.exclude_bboxes:
            return False
        
        # Use character center point for check
        char_center_x = (char_bbox[0] + char_bbox[2]) / 2
        char_center_y = (char_bbox[1] + char_bbox[3]) / 2
        
        for bbox in self.exclude_bboxes:
            # bbox = (x0, y0, x1, y1)
            if (bbox[0] <= char_center_x <= bbox[2] and 
                bbox[1] <= char_center_y <= bbox[3]):
                return True
        
        return False
    
    def _group_by_y(self, chars: List[Dict]) -> Dict[float, List[Dict]]:
        """Group characters by Y coordinate with tolerance."""
        lines_by_y = {}
        
        for char_info in chars:
            # Use origin Y if available, otherwise use bbox Y
            y = char_info['origin'][1] if char_info['origin'] else char_info['bbox'][1]
            
            # Find existing Y group within tolerance
            found_y = None
            for existing_y in lines_by_y.keys():
                if abs(existing_y - y) <= self.y_tolerance:
                    found_y = existing_y
                    break
            
            if found_y is None:
                found_y = y
                lines_by_y[found_y] = []
            
            lines_by_y[found_y].append(char_info)
        
        return lines_by_y
    
    def _build_lines(self, lines_by_y: Dict[float, List[Dict]]) -> List[str]:
        """Build text lines from character groups."""
        reconstructed_lines = []
        
        for y in sorted(lines_by_y.keys()):
            chars = lines_by_y[y]
            chars_sorted = sorted(chars, key=lambda c: c['bbox'][0])
            
            if not chars_sorted:
                continue
            
            # Build line text with appropriate spacing
            line_text = ""
            prev_x_end = None
            
            for char_info in chars_sorted:
                x_start = char_info['bbox'][0]
                char = char_info['c']
                
                if prev_x_end is not None:
                    gap = x_start - prev_x_end
                    # Add space if gap is significant
                    avg_char_width = char_info['size'] * 0.5
                    if gap > avg_char_width * 0.5:
                        line_text += " "
                
                line_text += char
                prev_x_end = char_info['bbox'][2]
            
            if line_text.strip():
                reconstructed_lines.append(line_text)
        
        return reconstructed_lines


# ============================================================================
# CJK Compatibility Character Mapping Function
# ============================================================================

def apply_cjk_compat_mapping(text: str) -> str:
    """
    Replace CJK Compatibility characters with their intended characters.
    
    These characters appear when Word documents are converted to PDF
    and font encoding is not properly preserved.
    
    Args:
        text: Text containing CJK Compatibility characters
        
    Returns:
        Text with characters replaced
    """
    if not text:
        return text
    
    result = text
    for cjk_char, replacement in CJK_COMPAT_CHAR_MAP.items():
        result = result.replace(cjk_char, replacement)
    
    return result


# ============================================================================
# Export
# ============================================================================

__all__ = [
    'TextQualityConfig',
    'TextQualityResult',
    'PageTextAnalysis',
    'TextQualityAnalyzer',
    'PageOCRFallbackEngine',
    'QualityAwareTextExtractor',
    'FragmentedTextReconstructor',
    'apply_cjk_compat_mapping',
    'CJK_COMPAT_CHAR_MAP',
]
