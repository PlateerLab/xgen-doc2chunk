# hwpx_helper/hwpx_section.py
"""
HWPX 섹션 파싱

HWPX 문서의 섹션 XML을 파싱하여 텍스트, 테이블, 이미지, 차트를 추출합니다.

테이블 처리:
- HWPXTableExtractor: hp:tbl 요소 → TableData 변환
- HWPXTableProcessor: TableData → HTML/Markdown/Text 출력

차트 처리:
- hp:chart 요소 발견 시 chart_callback 호출
- 원본 문서 순서대로 차트가 삽입됨
"""
import logging
import xml.etree.ElementTree as ET
import zipfile
from typing import Dict, Set, Optional, Callable

from xgen_doc2chunk.core.processor.hwpx_helper.hwpx_constants import HWPX_NAMESPACES
from xgen_doc2chunk.core.processor.hwpx_helper.hwpx_table_extractor import (
    HWPXTableExtractor,
)
from xgen_doc2chunk.core.processor.hwpx_helper.hwpx_table_processor import (
    HWPXTableProcessor,
)

from xgen_doc2chunk.core.functions.img_processor import ImageProcessor

logger = logging.getLogger("document-processor")

# Module-level instances (lazy initialized)
_table_extractor: Optional[HWPXTableExtractor] = None
_table_processor: Optional[HWPXTableProcessor] = None


def _get_table_extractor() -> HWPXTableExtractor:
    """Get or create the module-level table extractor."""
    global _table_extractor
    if _table_extractor is None:
        _table_extractor = HWPXTableExtractor()
    return _table_extractor


def _get_table_processor() -> HWPXTableProcessor:
    """Get or create the module-level table processor."""
    global _table_processor
    if _table_processor is None:
        _table_processor = HWPXTableProcessor()
    return _table_processor


def _process_table(
    table_element: ET.Element,
    ns: Dict[str, str],
    zf: zipfile.ZipFile = None,
    bin_item_map: Dict[str, str] = None,
    processed_images: Set[str] = None,
    image_processor: ImageProcessor = None,
) -> str:
    """Process a table element and return formatted output.
    
    Uses HWPXTableExtractor to convert XML to TableData,
    then HWPXTableProcessor to format as HTML.
    
    Cells holding a picture instead of text get an inline image tag so the
    OCR pass can read them in place. Without the zip/bin-item context the
    extractor falls back to its previous text-only behaviour.

    Args:
        table_element: hp:tbl XML element
        ns: Namespace dictionary
        zf: HWPX ZipFile, required to read image binaries
        bin_item_map: BinItem ID -> path mapping
        processed_images: Shared dedup set of already emitted image paths
        image_processor: Image processor used to save cell images

    Returns:
        Formatted table string (HTML)
    """
    extractor = _get_table_extractor()
    processor = _get_table_processor()

    if zf is not None and bin_item_map and image_processor is not None:
        def _resolve_cell_image(pic_elem: ET.Element) -> str:
            return _process_inline_image(
                pic_elem, zf, bin_item_map, processed_images, image_processor
            )
        extractor.configure_images(_resolve_cell_image)
    else:
        extractor.configure_images(None)

    try:
        table_data = extractor.extract_table(table_element, ns)
    finally:
        # Leave the shared extractor in its default state
        extractor.configure_images(None)

    if table_data:
        return processor.format_table(table_data)
    return ""


# Shape element tags that may contain drawText with text content
_SHAPE_TAGS = frozenset({'container', 'rect', 'polygon', 'ellipse', 'arc', 'curve'})


def _local_tag(elem: ET.Element) -> str:
    """Return the local tag name without namespace URI."""
    tag = elem.tag
    if '}' in tag:
        return tag.split('}', 1)[1]
    return tag


def _extract_sublist_text(
    sublist: ET.Element,
    ns: Dict[str, str],
    zf: zipfile.ZipFile = None,
    bin_item_map: Dict[str, str] = None,
    processed_images: Set[str] = None,
    image_processor: ImageProcessor = None,
    chart_callback: Optional[Callable[[str], str]] = None,
    section_headers: list = None,
    section_footers: list = None,
) -> str:
    """hp:subList > hp:p 요소에서 텍스트를 추출합니다.

    drawText (도형), header, footer, footnote, endnote 등에서 공통으로 사용됩니다.
    """
    parts = []
    for p in sublist.findall('hp:p', ns):
        p_parts = []
        for run in p.findall('hp:run', ns):
            run_parts = _process_run(
                run, ns, zf, bin_item_map, processed_images,
                image_processor, chart_callback, section_headers, section_footers
            )
            p_parts.extend(run_parts)
        if p_parts:
            parts.append("".join(p_parts))
    return "\n".join(parts)


def _extract_shape_text(
    shape_elem: ET.Element,
    ns: Dict[str, str],
    zf: zipfile.ZipFile = None,
    bin_item_map: Dict[str, str] = None,
    processed_images: Set[str] = None,
    image_processor: ImageProcessor = None,
    chart_callback: Optional[Callable[[str], str]] = None,
    section_headers: list = None,
    section_footers: list = None,
) -> str:
    """도형 요소(container, rect, polygon, ellipse 등)에서 텍스트를 재귀적으로 추출합니다.

    도형 내부 구조:
    - drawText > subList > p (텍스트 콘텐츠)
    - container 내부의 중첩 도형
    """
    texts = []
    tag = _local_tag(shape_elem)

    if tag == 'container':
        # container는 여러 도형을 그룹핑 - 중첩 도형을 재귀적으로 처리
        for child in shape_elem:
            child_tag = _local_tag(child)
            if child_tag in _SHAPE_TAGS:
                child_text = _extract_shape_text(
                    child, ns, zf, bin_item_map, processed_images,
                    image_processor, chart_callback, section_headers, section_footers
                )
                if child_text:
                    texts.append(child_text)

    # 이 도형의 drawText 콘텐츠 추출
    draw_text = shape_elem.find('hp:drawText', ns)
    if draw_text is not None:
        sub_list = draw_text.find('hp:subList', ns)
        if sub_list is not None:
            sl_text = _extract_sublist_text(
                sub_list, ns, zf, bin_item_map, processed_images,
                image_processor, chart_callback, section_headers, section_footers
            )
            if sl_text:
                texts.append(sl_text)

    return "\n".join(texts)


def _process_run(
    run: ET.Element,
    ns: Dict[str, str],
    zf: zipfile.ZipFile = None,
    bin_item_map: Dict[str, str] = None,
    processed_images: Set[str] = None,
    image_processor: ImageProcessor = None,
    chart_callback: Optional[Callable[[str], str]] = None,
    section_headers: list = None,
    section_footers: list = None,
) -> list:
    """hp:run 요소의 모든 자식을 문서 순서대로 처리합니다.

    처리 대상: 텍스트, 테이블, 차트, 이미지, 도형, ctrl(머리글/바닥글/각주 등).
    """
    parts = []

    for child in run:
        tag = _local_tag(child)

        if tag == 't':
            if child.text:
                parts.append(child.text)

        elif tag == 'tbl':
            table_html = _process_table(
                child, ns, zf, bin_item_map, processed_images, image_processor
            )
            if table_html:
                parts.append(f"\n{table_html}\n")

        elif tag == 'switch':
            case = child.find('hp:case', ns)
            if case is not None:
                chart = case.find('hp:chart', ns)
                if chart is not None and chart_callback:
                    chart_id_ref = chart.get('chartIDRef')
                    if chart_id_ref:
                        chart_text = chart_callback(chart_id_ref)
                        if chart_text:
                            parts.append(f"\n{chart_text}\n")

        elif tag == 'pic':
            if zf and bin_item_map:
                image_text = _process_inline_image(
                    child, zf, bin_item_map, processed_images, image_processor
                )
                if image_text:
                    parts.append(image_text)

        elif tag == 'ctrl':
            ctrl_parts = _process_ctrl(
                child, ns, zf, bin_item_map, processed_images,
                image_processor, chart_callback, section_headers, section_footers
            )
            parts.extend(ctrl_parts)

        elif tag in _SHAPE_TAGS:
            shape_text = _extract_shape_text(
                child, ns, zf, bin_item_map, processed_images,
                image_processor, chart_callback, section_headers, section_footers
            )
            if shape_text:
                parts.append(shape_text)

    return parts


def _process_ctrl(
    ctrl: ET.Element,
    ns: Dict[str, str],
    zf: zipfile.ZipFile = None,
    bin_item_map: Dict[str, str] = None,
    processed_images: Set[str] = None,
    image_processor: ImageProcessor = None,
    chart_callback: Optional[Callable[[str], str]] = None,
    section_headers: list = None,
    section_footers: list = None,
) -> list:
    """ctrl 요소의 자식을 처리합니다.

    ctrl 내부: header, footer, footNote, endNote, table, image, 도형 등.
    header/footer 텍스트는 섹션 레벨 리스트에 수집하여 출력 상단/하단에 배치합니다.
    """
    parts = []

    for child in ctrl:
        tag = _local_tag(child)

        if tag == 'tbl':
            table_html = _process_table(
                child, ns, zf, bin_item_map, processed_images, image_processor
            )
            if table_html:
                parts.append(f"\n{table_html}\n")

        elif tag == 'pic':
            # hp:pic 또는 hc:pic 모두 동일하게 처리
            if zf and bin_item_map:
                image_text = _process_inline_image(
                    child, zf, bin_item_map, processed_images, image_processor
                )
                if image_text:
                    parts.append(image_text)

        elif tag == 'header':
            if section_headers is not None:
                sub_list = child.find('hp:subList', ns)
                if sub_list is not None:
                    h_text = _extract_sublist_text(
                        sub_list, ns, zf, bin_item_map, processed_images,
                        image_processor, chart_callback, section_headers, section_footers
                    )
                    if h_text:
                        section_headers.append(f"[Header]\n{h_text}")

        elif tag == 'footer':
            if section_footers is not None:
                sub_list = child.find('hp:subList', ns)
                if sub_list is not None:
                    f_text = _extract_sublist_text(
                        sub_list, ns, zf, bin_item_map, processed_images,
                        image_processor, chart_callback, section_headers, section_footers
                    )
                    if f_text:
                        section_footers.append(f"[Footer]\n{f_text}")

        elif tag in ('footNote', 'endNote'):
            sub_list = child.find('hp:subList', ns)
            if sub_list is not None:
                note_text = _extract_sublist_text(
                    sub_list, ns, zf, bin_item_map, processed_images,
                    image_processor, chart_callback, section_headers, section_footers
                )
                if note_text:
                    label = "Footnote" if tag == 'footNote' else "Endnote"
                    parts.append(f"\n[{label}]\n{note_text}")

        elif tag in _SHAPE_TAGS:
            shape_text = _extract_shape_text(
                child, ns, zf, bin_item_map, processed_images,
                image_processor, chart_callback, section_headers, section_footers
            )
            if shape_text:
                parts.append(shape_text)

    return parts


def parse_hwpx_section(
    xml_content: bytes,
    zf: zipfile.ZipFile = None,
    bin_item_map: Dict[str, str] = None,
    processed_images: Set[str] = None,
    image_processor: ImageProcessor = None,
    chart_callback: Optional[Callable[[str], str]] = None
) -> str:
    """
    HWPX 섹션 XML을 파싱합니다.

    문단, 테이블, 인라인 이미지, 차트, 도형, 머리글/바닥글을 원본 문서 순서대로 처리합니다.

    HWPX structure:
    - <hs:sec> -> <hp:p> (최상위 문단)
    - <hp:p> -> <hp:run> -> <hp:t> (Text)
    - <hp:p> -> <hp:run> -> <hp:tbl> (Table)
    - <hp:p> -> <hp:run> -> <hp:ctrl> -> <hc:pic> (Image)
    - <hp:p> -> <hp:run> -> <hp:ctrl> -> <hp:header/footer> (Header/Footer)
    - <hp:p> -> <hp:run> -> <hp:switch> -> <hp:case> -> <hp:chart> (Chart)
    - <hp:p> -> <hp:run> -> <hp:pic> (Direct Image)
    - <hp:p> -> <hp:run> -> <hp:container/rect/polygon/ellipse> (Shapes with drawText)

    Args:
        xml_content: 섹션 XML 바이너리 데이터
        zf: ZipFile 객체 (이미지 추출용)
        bin_item_map: BinItem ID -> 파일 경로 매핑
        processed_images: 처리된 이미지 경로 집합 (중복 방지)
        image_processor: 이미지 프로세서 인스턴스
        chart_callback: 차트 참조 발견 시 호출할 콜백 함수
                       chartIDRef (예: "Chart/chart1.xml")를 받아 포맷된 차트 텍스트 반환

    Returns:
        추출된 텍스트 문자열
    """
    try:
        root = ET.fromstring(xml_content)
        ns = HWPX_NAMESPACES

        section_headers = []
        section_footers = []
        text_parts = []

        # 최상위 레벨의 hp:p만 처리 (테이블 내부의 hp:p는 테이블 파서에서 처리)
        for p in root.findall('hp:p', ns):
            p_text = []
            for run in p.findall('hp:run', ns):
                run_parts = _process_run(
                    run, ns, zf, bin_item_map, processed_images,
                    image_processor, chart_callback, section_headers, section_footers
                )
                p_text.extend(run_parts)

            if p_text:
                text_parts.append("".join(p_text))

        # 머리글을 상단에, 바닥글을 하단에 배치
        result = []
        result.extend(section_headers)
        result.extend(text_parts)
        result.extend(section_footers)

        return "\n".join(result)

    except Exception as e:
        logger.error(f"Error parsing HWPX XML: {e}")
        return ""


def _process_inline_image(
    pic: ET.Element,
    zf: zipfile.ZipFile,
    bin_item_map: Dict[str, str],
    processed_images: Optional[Set[str]],
    image_processor: ImageProcessor
) -> str:
    """
    인라인 이미지를 처리합니다.

    HWPX 이미지 구조:
    - <hp:pic> 또는 <hc:pic>
      - <hc:img binaryItemIDRef="image3">
    
    Args:
        pic: hp:pic 또는 hc:pic 요소
        zf: ZipFile 객체
        bin_item_map: BinItem ID -> 파일 경로 매핑
        processed_images: 처리된 이미지 경로 집합
        image_processor: 이미지 프로세서 인스턴스

    Returns:
        이미지 태그 문자열 또는 빈 문자열
    """
    ns = HWPX_NAMESPACES
    
    try:
        # Try to find binaryItemIDRef from nested hc:img element
        img_elem = pic.find('hc:img', ns)
        if img_elem is not None:
            bin_item_id = img_elem.get('binaryItemIDRef')
        else:
            # Fallback: try direct BinItem attribute
            bin_item_id = pic.get('BinItem')
        
        if not bin_item_id or bin_item_id not in bin_item_map:
            return ""

        img_path = bin_item_map[bin_item_id]

        # HWPX href might be relative. Usually "BinData/xxx.png"
        full_path = img_path
        if full_path not in zf.namelist():
            if f"Contents/{img_path}" in zf.namelist():
                full_path = f"Contents/{img_path}"

        if full_path not in zf.namelist():
            return ""

        with zf.open(full_path) as f:
            image_data = f.read()

        image_tag = image_processor.save_image(image_data)
        if image_tag:
            if processed_images is not None:
                processed_images.add(full_path)
            return f"\n{image_tag}\n"

    except Exception as e:
        logger.warning(f"Failed to process inline image: {e}")

    return ""
