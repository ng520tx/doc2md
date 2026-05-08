"""
文档转 Markdown 工具 (doc2md)
将 .doc / .docx / .pdf 文档转换为 Markdown + 图片，供 AI 完整阅读需求文档。

使用方式:
    python doc2md.py <文件路径>                    # 输出到同目录
    python doc2md.py <文件路径> -o <输出目录>       # 指定输出目录
    python doc2md.py <文件路径> --embed-images      # 图片以 base64 内嵌到 md 中

输出:
    <文件名>.md          - Markdown 文本
    <文件名>_images/     - 提取的图片文件夹
"""

import argparse
import base64
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

AI_VIEWABLE_EXTS = {".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp"}
CONVERTIBLE_EXTS = {".emf", ".wmf", ".tiff", ".tif", ".svg"}

PDF_ENGINES = ["auto", "mineru", "docling", "pymupdf4llm"]
PDF_ENGINE_LABELS = {
    "auto": "自动选择 (Auto)",
    "mineru": "MinerU (中文最佳, 表格 0.87)",
    "docling": "Docling (综合最佳, 表格 0.89)",
    "pymupdf4llm": "PyMuPDF4LLM (最快, 表格 0.40)",
}


# ─── DOCX 转换 ────────────────────────────────────────────────────────────────

def convert_docx(input_path: Path, output_dir: Path, embed_images: bool = False,
                 progress_cb=None) -> Path:
    try:
        import mammoth
        from markdownify import markdownify as md_convert
        from docx import Document
    except ImportError as e:
        raise RuntimeError("Missing dependencies: pip install mammoth markdownify python-docx Pillow") from e

    stem = input_path.stem
    images_dir = output_dir / f"{stem}_images"
    images_dir.mkdir(parents=True, exist_ok=True)

    if progress_cb:
        progress_cb(10, "mammoth HTML...")

    image_counter = [0]

    def handle_image(image):
        image_counter[0] += 1
        ext = _guess_image_ext(image.content_type)
        filename = f"img_{image_counter[0]:03d}{ext}"
        filepath = images_dir / filename
        with image.open() as img_stream:
            img_bytes = img_stream.read()
        with open(filepath, "wb") as f:
            f.write(img_bytes)
        if embed_images:
            b64 = base64.b64encode(img_bytes).decode("utf-8")
            return {"src": f"data:{image.content_type};base64,{b64}"}
        return {"src": f"{stem}_images/{filename}"}

    try:
        with open(input_path, "rb") as docx_file:
            result = mammoth.convert_to_html(
                docx_file,
                convert_image=mammoth.images.img_element(handle_image),
            )
    except Exception as e:
        raise RuntimeError(f"DOCX HTML conversion failed ({input_path.name}): {e}") from e

    if progress_cb:
        progress_cb(30, "Markdown...")

    html = result.value
    if result.messages:
        for msg in result.messages:
            print(f"  [WARN] mammoth: {msg}")
    markdown_text = md_convert(html, heading_style="ATX", strip=["script", "style"])

    _extract_extra_images_from_zip(input_path, images_dir, image_counter[0])

    if progress_cb:
        progress_cb(50, "structure...")

    doc = Document(str(input_path))
    structure = _extract_doc_structure(doc)

    markdown_text = _remove_raw_toc(markdown_text, structure)
    markdown_text = _apply_structure(markdown_text, structure)
    markdown_text = _fix_ordered_lists(markdown_text)
    markdown_text = _fix_table_headers(markdown_text)
    markdown_text = _fix_merged_cells(markdown_text)
    markdown_text = _extract_footnotes(doc, markdown_text)
    markdown_text = _clean_markdown(markdown_text)

    if progress_cb:
        progress_cb(70, "images...")

    converted_map = _convert_special_images(images_dir)
    _remove_empty_files(images_dir)

    if converted_map:
        for old_name, new_name in converted_map.items():
            markdown_text = markdown_text.replace(f"{stem}_images/{old_name}", f"{stem}_images/{new_name}")
            markdown_text = markdown_text.replace(old_name, new_name)

    header = _build_header(input_path, "DOCX", len(list(images_dir.glob("*"))))
    markdown_text = header + markdown_text

    if progress_cb:
        progress_cb(90, "saving...")

    md_path = output_dir / f"{stem}.md"
    md_path.write_text(markdown_text, encoding="utf-8")

    img_count = len(list(images_dir.glob("*")))
    print(f"[OK] DOCX -> {md_path}")
    print(f"  images: {img_count}, dir: {images_dir}")

    if progress_cb:
        progress_cb(100, "done")

    return md_path


def _extract_doc_structure(doc) -> dict:
    """提取文档标题结构。
    优先级：
    1. TOC 段落（toc 1/2/3/4 样式）有完整 level + 编号 + 文本，作为权威锚点；
       同时扫描全部段落，借助已对齐的 (numId, ilvl) → level 映射推断 TOC 之外
       的更深层级标题（5 级及以上），并按上下文生成 full_number。
    2. 无 TOC 时回退：Heading X 样式 + numbering 计算编号。
    """
    toc_items = _extract_toc_items(doc)

    if toc_items:
        return _build_structure_from_toc(doc, toc_items)

    headings = _extract_headings_by_style(doc)
    return {"headings": headings, "toc": [], "source": "style"}


_TOC_ALIGN_LOOKAHEAD = 80


def _build_structure_from_toc(doc, toc_items: list[dict]) -> dict:
    """以 TOC 为锚点，结合段落 numId+ilvl 推断更深层级标题，构建完整 headings。"""
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

    candidates = _collect_title_candidates(doc, ns)
    _align_candidates_to_toc(candidates, toc_items)
    ilvl_to_level, heading_num_ids = _learn_ilvl_to_level(candidates)
    headings = _build_headings_with_inference(
        candidates, ilvl_to_level, heading_num_ids, toc_items
    )

    return {"headings": headings, "toc": toc_items, "source": "toc+infer"}


def _collect_title_candidates(doc, ns: dict) -> list[dict]:
    """扫描所有段落，收集候选标题段落。
    候选条件：使用 Heading X 样式，或拥有 numId+ilvl（不论来自段落本身还是样式继承）。
    跳过 TOC 段落本身（toc 1/2/3/4 样式）。"""
    candidates = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        style_name = para.style.name if para.style else ""

        if re.match(r"toc\s*\d", style_name, re.IGNORECASE):
            continue

        num_id, ilvl = _get_para_num_info(para, ns)
        if num_id is None and ilvl is None:
            num_id, ilvl = _get_style_num_info(para.style, ns)

        is_heading_style = bool(re.match(r"[Hh]eading\s*\d", style_name))
        has_numbering = num_id is not None and ilvl is not None
        if not is_heading_style and not has_numbering:
            continue

        candidates.append({
            "text": text,
            "style": style_name,
            "num_id": num_id,
            "ilvl": ilvl,
            "is_heading_style": is_heading_style,
            "matched_toc": None,
        })
    return candidates


def _align_candidates_to_toc(candidates: list[dict], toc_items: list[dict]) -> None:
    """按顺序把 TOC 项与候选段落对齐，给匹配项写入 matched_toc。
    空 clean_text 的 toc 项（如 '5.2.' 占位空标题）跳过对齐——
    它们由 _build_headings_with_inference 按 toc 顺序单独补齐。"""
    cursor = 0
    for toc in toc_items:
        target = toc["clean_text"]
        if not target:
            continue
        end = min(len(candidates), cursor + _TOC_ALIGN_LOOKAHEAD)
        for i in range(cursor, end):
            tc = candidates[i]
            if tc["matched_toc"] is not None:
                continue
            if _text_similar(tc["text"], target):
                tc["matched_toc"] = toc
                cursor = i + 1
                break


def _learn_ilvl_to_level(candidates: list[dict]) -> tuple[dict, set]:
    """从已对齐 TOC 的候选中，学习 (num_id, ilvl) → level 映射。
    同时记录哪些 num_id 属于"标题列表"，避免把普通有序列表项当作标题。"""
    mapping: dict = {}
    heading_num_ids: set = set()
    for tc in candidates:
        if tc["matched_toc"] is None:
            continue
        if tc["num_id"] is None or tc["ilvl"] is None:
            continue
        key = (tc["num_id"], tc["ilvl"])
        mapping[key] = tc["matched_toc"]["level"]
        heading_num_ids.add(tc["num_id"])
    return mapping, heading_num_ids


def _build_headings_with_inference(candidates: list[dict],
                                    ilvl_to_level: dict,
                                    heading_num_ids: set,
                                    toc_items: list[dict] | None = None) -> list[dict]:
    """按文档顺序构建完整 headings 清单。
    - TOC 已匹配段落：直接采用 TOC 的 level/full_number/clean_text
    - 未匹配但属于"标题列表"段落：按映射或同 num_id 的相邻 ilvl 推断 level，
      并基于父级 prefix 与本级计数器生成 full_number
    - 不属于标题列表的段落：当作普通列表项跳过
    - 未对齐到候选的 toc 项（如 '5.2.' 空标题）：按 toc 顺序补齐到合适位置
    """
    by_num_id: dict = {}
    for (num_id, ilvl), level in ilvl_to_level.items():
        by_num_id.setdefault(num_id, {})[ilvl] = level

    toc_items = toc_items or []
    toc_idx_map = {id(toc): i for i, toc in enumerate(toc_items)}

    headings: list[dict] = []
    counters: dict = {}
    prefix_at_level: dict = {}
    next_toc_idx = 0

    def _emit(level: int, full_number: str, clean_text: str,
              from_toc: bool) -> None:
        for k in list(counters.keys()):
            if k > level:
                del counters[k]
        for k in list(prefix_at_level.keys()):
            if k > level:
                del prefix_at_level[k]

        if from_toc:
            nums = re.findall(r"\d+", full_number)
            if nums:
                counters[level] = int(nums[-1])
        prefix_at_level[level] = full_number.rstrip(".")

        headings.append({
            "level": level,
            "text": _join_number_text(full_number, clean_text),
            "full_number": full_number,
            "clean_text": clean_text,
        })

    def _flush_toc_until(target_idx: int) -> None:
        nonlocal next_toc_idx
        while next_toc_idx < target_idx:
            toc = toc_items[next_toc_idx]
            _emit(toc["level"], toc["full_number"], toc["clean_text"], True)
            next_toc_idx += 1

    for tc in candidates:
        level = _resolve_level(tc, ilvl_to_level, by_num_id, heading_num_ids)
        if level is None:
            continue

        if tc["matched_toc"] is not None:
            toc = tc["matched_toc"]
            t_idx = toc_idx_map.get(id(toc), next_toc_idx)
            if t_idx > next_toc_idx:
                _flush_toc_until(t_idx)
            _emit(toc["level"], toc["full_number"], toc["clean_text"], True)
            next_toc_idx = t_idx + 1
        else:
            counters[level] = counters.get(level, 0) + 1
            cnt = counters[level]
            parent_prefix = ""
            for k in range(level - 1, 0, -1):
                if prefix_at_level.get(k):
                    parent_prefix = prefix_at_level[k]
                    break
            full_number = f"{parent_prefix}.{cnt}." if parent_prefix else f"{cnt}."
            _emit(level, full_number, tc["text"], False)

    _flush_toc_until(len(toc_items))

    return headings


def _resolve_level(tc: dict, ilvl_to_level: dict,
                   by_num_id: dict, heading_num_ids: set) -> int | None:
    """确定一个候选段落的标题层级。返回 None 表示这不是标题。"""
    if tc["matched_toc"] is not None:
        return tc["matched_toc"]["level"]

    if tc["is_heading_style"]:
        m = re.match(r"[Hh]eading\s*(\d)", tc["style"])
        if m:
            return int(m.group(1))

    if tc["num_id"] is None or tc["ilvl"] is None:
        return None
    if tc["num_id"] not in heading_num_ids:
        return None

    key = (tc["num_id"], tc["ilvl"])
    if key in ilvl_to_level:
        return ilvl_to_level[key]

    same_num = by_num_id.get(tc["num_id"], {})
    if not same_num:
        return None
    max_known_ilvl = max(same_num.keys())
    max_known_level = max(same_num.values())
    inferred = max_known_level + (tc["ilvl"] - max_known_ilvl)
    return max(1, min(6, inferred))


def _extract_toc_items(doc) -> list[dict]:
    """从 docx 中提取 TOC 项。每项含 level/full_number/clean_text。"""
    items = []
    for para in doc.paragraphs:
        style_name = para.style.name if para.style else ""
        text = para.text.strip()

        toc_match = re.match(r"toc\s*(\d)", style_name, re.IGNORECASE)
        if not toc_match or not text:
            continue

        level = int(toc_match.group(1))
        raw = re.sub(r"\t.*$", "", text).strip()
        raw = re.sub(r"\s*\d+\s*$", "", raw).strip()
        if not raw:
            continue

        full_number, clean_text = _split_number_text(raw)
        if not full_number and not clean_text:
            continue

        items.append({
            "level": level,
            "full_number": full_number,
            "clean_text": clean_text,
        })
    return items


def _split_number_text(text: str) -> tuple[str, str]:
    """把 '1.1. 总体描述' 拆成 ('1.1.', '总体描述')。
    支持 '5.2.' 这种纯编号无文本（空标题）→ ('5.2.', '')。"""
    m = re.match(r"^(\d+(?:\.\d+)*\.?)\s+(.*)$", text)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    m = re.match(r"^(\d+(?:\.\d+)*\.?)\s*$", text)
    if m:
        return m.group(1).strip(), ""
    return "", text.strip()


def _join_number_text(full_number: str, clean_text: str) -> str:
    if full_number:
        return f"{full_number} {clean_text}".strip()
    return clean_text


def _extract_headings_by_style(doc) -> list[dict]:
    """无 TOC 时的回退方案：识别 Heading X 样式，配合 numbering 计算编号。"""
    heading_texts = []
    numbering_info = _parse_numbering(doc)
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    counters: dict = {}
    heading_num_id = None

    for para in doc.paragraphs:
        style_name = para.style.name if para.style else ""
        text = para.text.strip()

        heading_match = re.match(r"[Hh]eading\s*(\d)", style_name)
        if heading_match and text:
            level = int(heading_match.group(1))

            num_id, ilvl = _get_para_num_info(para, ns)
            if num_id is None and ilvl is None:
                num_id, ilvl = _get_style_num_info(para.style, ns)
            if ilvl is None:
                ilvl = level - 1

            if num_id is not None and numbering_info:
                if heading_num_id is None:
                    heading_num_id = num_id

                if num_id == heading_num_id:
                    numbered_text = _compute_heading_number(
                        numbering_info, num_id, ilvl, counters, text
                    )
                else:
                    numbered_text = text
            else:
                numbered_text = text

            full_number, clean_text = _split_number_text(numbered_text)
            if not clean_text:
                clean_text = text

            heading_texts.append({
                "level": level,
                "text": numbered_text,
                "full_number": full_number,
                "clean_text": clean_text,
            })
    return heading_texts


def _parse_numbering(doc) -> dict | None:
    """解析 numbering.xml，构建完整的编号定义。
    支持：abstractNum、num（含 lvlOverride/startOverride）、lvlRestart、多种 numFmt。"""
    from lxml import etree
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

    numbering_part = None
    for rel in doc.part.rels.values():
        if "numbering" in rel.reltype:
            numbering_part = rel.target_part
            break
    if not numbering_part:
        return None

    root = etree.fromstring(numbering_part.blob)

    abstract_defs = {}
    for an in root.findall(".//w:abstractNum", ns):
        an_id = an.get(f'{{{ns["w"]}}}abstractNumId')
        levels = {}
        for lvl in an.findall(".//w:lvl", ns):
            levels[int(lvl.get(f'{{{ns["w"]}}}ilvl'))] = _parse_lvl(lvl, ns)
        abstract_defs[an_id] = levels

    num_to_abstract = {}
    num_overrides = {}
    for num in root.findall(".//w:num", ns):
        num_id = num.get(f'{{{ns["w"]}}}numId')
        abs_ref = num.find("w:abstractNumId", ns)
        if abs_ref is not None:
            num_to_abstract[num_id] = abs_ref.get(f'{{{ns["w"]}}}val')
        overrides = {}
        for ovr in num.findall("w:lvlOverride", ns):
            ovr_ilvl = int(ovr.get(f'{{{ns["w"]}}}ilvl'))
            start_el = ovr.find("w:startOverride", ns)
            lvl_el = ovr.find("w:lvl", ns)
            entry = {}
            if start_el is not None:
                entry["start"] = int(start_el.get(f'{{{ns["w"]}}}val'))
            if lvl_el is not None:
                entry["lvl"] = _parse_lvl(lvl_el, ns)
            overrides[ovr_ilvl] = entry
        if overrides:
            num_overrides[num_id] = overrides

    return {
        "abstract": abstract_defs,
        "num_map": num_to_abstract,
        "overrides": num_overrides,
    }


def _parse_lvl(lvl_el, ns: dict) -> dict:
    """解析单个 w:lvl 元素，提取 fmt、text、start、restart"""
    fmt_el = lvl_el.find("w:numFmt", ns)
    fmt = fmt_el.get(f'{{{ns["w"]}}}val') if fmt_el is not None else "none"
    text_el = lvl_el.find("w:lvlText", ns)
    text = text_el.get(f'{{{ns["w"]}}}val') if text_el is not None else ""
    start_el = lvl_el.find("w:start", ns)
    start = int(start_el.get(f'{{{ns["w"]}}}val')) if start_el is not None else 1
    restart_el = lvl_el.find("w:lvlRestart", ns)
    restart = int(restart_el.get(f'{{{ns["w"]}}}val')) if restart_el is not None else None
    return {"fmt": fmt, "text": text, "start": start, "restart": restart}


def _get_para_num_info(para, ns: dict) -> tuple:
    """从段落 XML 直接获取 numId 和 ilvl"""
    pPr = para._element.find("w:pPr", ns)
    if pPr is None:
        return None, None
    numPr = pPr.find("w:numPr", ns)
    if numPr is None:
        return None, None
    numId_el = numPr.find("w:numId", ns)
    ilvl_el = numPr.find("w:ilvl", ns)
    num_id = numId_el.get(f'{{{ns["w"]}}}val') if numId_el is not None else None
    ilvl = int(ilvl_el.get(f'{{{ns["w"]}}}val')) if ilvl_el is not None else None
    return num_id, ilvl


def _get_style_num_info(style, ns: dict) -> tuple:
    """从样式定义中获取 numId 和 ilvl，沿 basedOn 继承链向上查找"""
    visited = set()
    current = style
    while current is not None and current.element is not None:
        style_id = current.style_id
        if style_id in visited:
            break
        visited.add(style_id)

        sPr = current.element.find("w:pPr", ns)
        if sPr is not None:
            numPr = sPr.find("w:numPr", ns)
            if numPr is not None:
                numId_el = numPr.find("w:numId", ns)
                ilvl_el = numPr.find("w:ilvl", ns)
                num_id = numId_el.get(f'{{{ns["w"]}}}val') if numId_el is not None else None
                ilvl = int(ilvl_el.get(f'{{{ns["w"]}}}val')) if ilvl_el is not None else None
                if num_id is not None:
                    return num_id, ilvl

        try:
            current = current.base_style
        except Exception:
            break
    return None, None


_NUM_FMT_FUNCS = {
    "decimal": lambda n: str(n),
    "lowerLetter": lambda n: chr(ord('a') + (n - 1) % 26) if n > 0 else "",
    "upperLetter": lambda n: chr(ord('A') + (n - 1) % 26) if n > 0 else "",
    "lowerRoman": lambda n: _to_roman(n).lower(),
    "upperRoman": lambda n: _to_roman(n),
    "chineseCounting": lambda n: _to_chinese(n),
    "chineseCountingThousand": lambda n: _to_chinese(n),
    "ideographTraditional": lambda n: _to_chinese(n),
}


def _to_roman(n: int) -> str:
    vals = [(1000, 'M'), (900, 'CM'), (500, 'D'), (400, 'CD'), (100, 'C'),
            (90, 'XC'), (50, 'L'), (40, 'XL'), (10, 'X'), (9, 'IX'),
            (5, 'V'), (4, 'IV'), (1, 'I')]
    result = ""
    for v, s in vals:
        while n >= v:
            result += s
            n -= v
    return result


def _to_chinese(n: int) -> str:
    digits = "零一二三四五六七八九十"
    if 0 <= n <= 10:
        return digits[n]
    if n < 20:
        return f"十{digits[n - 10]}" if n > 10 else "十"
    return str(n)


def _format_num(value: int, fmt: str) -> str:
    func = _NUM_FMT_FUNCS.get(fmt)
    if func:
        return func(value)
    return str(value)


def _compute_heading_number(numbering_info: dict, num_id: str, ilvl: int,
                            counters: dict, text: str) -> str:
    """根据编号定义和计数器状态，计算标题的完整编号并拼接到文本前。
    支持 lvlOverride/startOverride、lvlRestart、多种 numFmt。"""
    if re.match(r"^\d+(\.\d+)*\.?\s", text):
        return text

    abs_id = numbering_info["num_map"].get(num_id)
    if abs_id is None:
        return text
    levels = numbering_info["abstract"].get(abs_id, {})
    lvl_def = dict(levels.get(ilvl, {}))
    if not lvl_def or lvl_def.get("fmt") in ("none", "bullet"):
        return text

    overrides = numbering_info.get("overrides", {}).get(num_id, {})
    if ilvl in overrides:
        ovr = overrides[ilvl]
        if "lvl" in ovr:
            lvl_def.update(ovr["lvl"])
        if "start" in ovr and ilvl not in counters:
            counters[ilvl] = ovr["start"] - 1

    restart_lvl = lvl_def.get("restart")
    if restart_lvl is not None and restart_lvl > 0:
        pass

    counters[ilvl] = counters.get(ilvl, lvl_def.get("start", 1) - 1) + 1
    for deeper in list(counters.keys()):
        if deeper > ilvl:
            del counters[deeper]

    tmpl = lvl_def.get("text", "")
    result = tmpl
    for lvl_idx in range(ilvl + 1):
        placeholder = f"%{lvl_idx + 1}"
        val = counters.get(lvl_idx, 0)
        lvl_fmt = levels.get(lvl_idx, {}).get("fmt", "decimal")
        result = result.replace(placeholder, _format_num(val, lvl_fmt))

    return f"{result} {text}"


def _remove_raw_toc(markdown_text: str, structure: dict) -> str:
    """移除 markdown 中由 mammoth 输出的原始 TOC 段落（toc 1/2/3/4 内容）。"""
    if not structure.get("toc"):
        return markdown_text

    toc_texts = set()
    for item in structure["toc"]:
        full = _join_number_text(item.get("full_number", ""), item.get("clean_text", ""))
        if full:
            toc_texts.add(full)
        if item.get("clean_text"):
            toc_texts.add(item["clean_text"])

    lines = markdown_text.split("\n")
    result = []
    skip_count = 0

    for line in lines:
        stripped = line.strip()
        clean_line = re.sub(r"\s*\d+\s*$", "", stripped).strip()

        if clean_line in toc_texts or stripped in toc_texts:
            skip_count += 1
            if skip_count <= len(structure["toc"]) + 2:
                continue

        if skip_count > 0 and stripped == "":
            continue

        if stripped:
            skip_count = 0

        result.append(line)

    return "\n".join(result)


_HEADING_LOOKAHEAD = 3


def _apply_structure(markdown_text: str, structure: dict) -> str:
    """按 TOC/Headings 顺序，在 markdown 中定位"伪标题"行并替换为标准 # 形式。

    支持的伪标题形式（来自 mammoth 对 docx 不同标题样式的转换结果）：
      a) ``# / ## / ### ...`` 已是标题
      b) ``1. **需求说明**`` 顶层有序列表 + 加粗（H1）
      c) ``* + - 1. 启动页`` 嵌套列表项（多级标题）
      d) ``   3. APP客户端`` 缩进有序列表（多级标题）
      e) ``**需求说明**`` 单独加粗段落

    对空 clean_text 的"占位标题"（如 docx 中 '5.2.' 没有文本内容），
    它们在 mammoth 输出中没有对应可见行，会在下一个匹配标题之前主动插入；
    若是末尾标题则追加在文档末尾。
    """
    headings = structure.get("headings", [])

    toc_block = _build_toc_block(headings)

    lines = markdown_text.split("\n")
    result: list[str] = []
    h_idx = 0
    toc_inserted = False

    def _flush_empty_headings_until(target_idx: int) -> None:
        """h_idx 推进到 target_idx 之前，把中间未匹配的空 clean_text heading 输出。"""
        nonlocal h_idx
        while h_idx < target_idx:
            h = headings[h_idx]
            if not (h.get("clean_text") or "").strip():
                result.append("")
                result.append(_build_heading_line(h))
                result.append("")
            h_idx += 1

    for line in lines:
        stripped = line.strip()

        if not toc_inserted and toc_block and (
            stripped == "\u76ee\u5f55"
            or stripped == "**\u76ee\u5f55**"
            or stripped == "**\u76ee \u5f55**"
        ):
            result.append(toc_block)
            toc_inserted = True
            continue

        if h_idx < len(headings):
            cand = _extract_heading_candidate_text(line)
            if cand:
                match_idx = _find_matching_heading(cand, headings, h_idx, _HEADING_LOOKAHEAD)
                if match_idx >= 0:
                    _flush_empty_headings_until(match_idx)
                    target = headings[match_idx]
                    result.append("")
                    result.append(_build_heading_line(target))
                    result.append("")
                    h_idx = match_idx + 1
                    continue

        result.append(line)

    _flush_empty_headings_until(len(headings))

    if toc_block and not toc_inserted:
        result.insert(0, toc_block)

    return "\n".join(result)


def _build_toc_block(headings: list[dict]) -> str:
    """基于完整 headings 清单（含推断出的深层标题）生成文档目录块。"""
    if not headings:
        return ""
    block = "\n---\n**[[ 文档目录 ]]**\n\n"
    for item in headings:
        level = max(1, min(6, int(item.get("level", 1))))
        indent = "  " * (level - 1)
        full = item.get("text") or _join_number_text(
            item.get("full_number", ""), item.get("clean_text", "")
        )
        block += f"{indent}- {full}\n"
    block += "---\n"
    return block


def _build_heading_line(target: dict) -> str:
    level = target["level"]
    full_number = target.get("full_number", "")
    if "clean_text" in target:
        clean_text = target.get("clean_text") or ""
    else:
        clean_text = target.get("text", "")
    if full_number and clean_text:
        return f"{'#' * level} {full_number} {clean_text}"
    if full_number:
        return f"{'#' * level} {full_number}"
    return f"{'#' * level} {clean_text}".rstrip()


def _find_matching_heading(candidate: str, headings: list[dict],
                           start: int, lookahead: int) -> int:
    """在 [start, start+lookahead) 范围内查找文本相似的 heading，返回索引或 -1。"""
    end = min(start + lookahead, len(headings))
    for i in range(start, end):
        if _text_similar(candidate, headings[i].get("clean_text", "")):
            return i
    return -1


_INLINE_DECOR_PATTERN = re.compile(r"(\*\*|__|~~|`)+")


def _strip_inline_decor(text: str) -> str:
    """去除 markdown 行内修饰符 (** __ ~~ `) 仅保留纯文本。"""
    text = _INLINE_DECOR_PATTERN.sub("", text)
    return text.strip()


def _strip_leading_number(text: str) -> str:
    """去除前导编号 ``1.``、``1.2.``、``1)`` 等。"""
    return re.sub(r"^\s*\d+(?:\.\d+)*[\.\)]?\s*", "", text)


def _extract_heading_candidate_text(line: str) -> str | None:
    """从 markdown 行中提取候选标题文本（清理掉编号和修饰符）。
    返回 None 表示不是标题候选。

    设计原则：宁可提取过多候选，由 _find_matching_heading 通过 TOC 文本匹配做最终过滤。
    """
    stripped = line.strip()
    if not stripped:
        return None

    m = re.match(r"^(#{1,6})\s+(.+?)\s*$", stripped)
    if m:
        text = _strip_inline_decor(m.group(2))
        text = _strip_leading_number(text).strip()
        return text or None

    m = re.match(r"^(\d+)\.\s+\*\*(.+?)\*\*\s*$", stripped)
    if m:
        return _strip_inline_decor(m.group(2)).strip() or None

    m = re.match(r"^[*+\-](?:\s+[*+\-])*\s+\d+\.\s+(.+?)\s*$", stripped)
    if m:
        text = _strip_inline_decor(m.group(1))
        text = _strip_leading_number(text).strip()
        return text or None

    if line and (line.startswith(" ") or line.startswith("\t")):
        m = re.match(r"^\s+(\d+)\.\s+(.+?)\s*$", line)
        if m:
            text = _strip_inline_decor(m.group(2))
            text = _strip_leading_number(text).strip()
            return text or None

    m = re.match(r"^(\d+)\.\s+(.+?)\s*$", stripped)
    if m:
        text = _strip_inline_decor(m.group(2))
        text = _strip_leading_number(text).strip()
        if text and not _looks_like_list_content(text):
            return text

    m = re.match(r"^\*\*(.+?)\*\*\s*$", stripped)
    if m:
        text = m.group(1).strip()
        if 1 <= len(text) <= 60 and not text.endswith(("：", ":", "。", ".")):
            text = _strip_leading_number(text).strip()
            return text or None

    if 2 <= len(stripped) <= 30:
        if not any(ch in stripped for ch in "。；;，,：:！!？?\u3002\uff1b\uff0c\uff1a\uff01\uff1f"):
            text = _strip_inline_decor(stripped)
            text = _strip_leading_number(text).strip()
            if text and re.search(r"[\u4e00-\u9fff\w]", text):
                return text

    return None


def _looks_like_list_content(text: str) -> bool:
    """判断文本是否看起来像普通列表内容而非标题。
    标题通常较短、不含句末标点、不含赋值符号等。"""
    if len(text) > 40:
        return True
    if any(ch in text for ch in "。；;"):
        return True
    if text.endswith(("，", ",", "：", ":")):
        return True
    if "=" in text and not re.search(r"[\u4e00-\u9fff]{3,}$", text):
        return True
    return False


_HEADING_TAIL_DELIM = "-—–_(（[【\\、:：/／"


def _text_similar(a: str, b: str) -> bool:
    """文本是否相似（忽略编号、空白、大小写）。
    支持候选文本以目标开头并紧跟分隔符的情况，例如：
      候选 "用户组织--功能没有"  目标 "用户组织"  -> 相似
      候选 "（一级）用户管理--已经有了" 目标 "（一级）用户管理" -> 相似
    """
    def normalize(s: str) -> str:
        s = re.sub(r"^\d+(\.\d+)*[\.\)]?\s*", "", s)
        s = re.sub(r"\s+", "", s)
        s = s.replace("\uff5e", "~")
        return s.lower()

    na, nb = normalize(a), normalize(b)
    if na == nb:
        return True

    if len(nb) >= 3 and len(na) > len(nb) and na.startswith(nb):
        next_ch = na[len(nb)]
        if next_ch in _HEADING_TAIL_DELIM:
            return True

    return False


def _fix_ordered_lists(text: str) -> str:
    lines = text.split("\n")
    result = []
    counter = 0
    in_list_context = False

    for line in lines:
        m = re.match(r"^(\s*)1\.\s+(.+)$", line)
        if m:
            indent = m.group(1)
            content = m.group(2)
            counter += 1
            result.append(f"{indent}{counter}. {content}")
            in_list_context = True
        else:
            stripped = line.strip()
            if in_list_context and (
                stripped == "" or stripped.startswith("![") or stripped.startswith("*")
            ):
                result.append(line)
            else:
                if not re.match(r"^\s*\d+\.\s", stripped):
                    counter = 0
                    in_list_context = False
                result.append(line)

    return "\n".join(result)


def _fix_table_headers(text: str) -> str:
    """修复表格首行为空的情况：如果第一行全空但第二行是分隔符、第三行有内容，
    则将第三行提升为表头"""
    lines = text.split("\n")
    result = []
    i = 0
    while i < len(lines):
        if (i + 2 < len(lines)
            and re.match(r"^\|\s*\|\s*(\|\s*)*$", lines[i].strip())
            and re.match(r"^\|\s*---", lines[i + 1].strip())
            and re.match(r"^\|", lines[i + 2].strip())):
            result.append(lines[i + 2])
            result.append(lines[i + 1])
            i += 3
        else:
            result.append(lines[i])
            i += 1
    return "\n".join(result)


def _fix_merged_cells(text: str) -> str:
    """修复表格中合并单元格导致的列数不一致问题"""
    lines = text.split("\n")
    result = []
    table_lines = []
    in_table = False

    def flush_table():
        if not table_lines:
            return
        col_counts = []
        for tl in table_lines:
            cols = tl.strip().split("|")
            if cols and cols[0] == "":
                cols = cols[1:]
            if cols and cols[-1].strip() == "":
                cols = cols[:-1]
            col_counts.append(len(cols))
        if not col_counts:
            result.extend(table_lines)
            return
        max_cols = max(col_counts)
        for tl in table_lines:
            cols = tl.strip().split("|")
            if cols and cols[0] == "":
                cols = cols[1:]
            if cols and cols[-1].strip() == "":
                cols = cols[:-1]
            while len(cols) < max_cols:
                cols.append(" ")
            is_sep = all(re.match(r"^\s*-{3,}\s*$", c) for c in cols)
            if is_sep:
                result.append("| " + " | ".join("---" for _ in cols) + " |")
            else:
                result.append("| " + " | ".join(c.strip() for c in cols) + " |")
        table_lines.clear()

    for line in lines:
        stripped = line.strip()
        if stripped.startswith("|") and stripped.endswith("|"):
            in_table = True
            table_lines.append(line)
        else:
            if in_table:
                flush_table()
                in_table = False
            result.append(line)

    if table_lines:
        flush_table()

    return "\n".join(result)


def _extract_footnotes(doc, text: str) -> str:
    """从 python-docx 提取脚注/尾注，追加到 Markdown 末尾"""
    footnotes = []
    try:
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        footnotes_part = None
        for rel in doc.part.rels.values():
            if "footnotes" in rel.reltype:
                footnotes_part = rel.target_part
                break
        if footnotes_part:
            from lxml import etree
            ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
            root = etree.fromstring(footnotes_part.blob)
            for fn in root.findall(".//w:footnote", ns):
                fn_id = fn.get(f"{{{ns['w']}}}id")
                if fn_id in ("-1", "0"):
                    continue
                texts = []
                for t in fn.findall(".//w:t", ns):
                    if t.text:
                        texts.append(t.text)
                if texts:
                    footnotes.append((fn_id, "".join(texts).strip()))
    except Exception as e:
        print(f"  [WARN] Footnote extraction failed: {e}")

    if not footnotes:
        return text

    text += "\n\n---\n\n### 脚注\n\n"
    for fn_id, fn_text in footnotes:
        text += f"[^{fn_id}]: {fn_text}\n"
        text = re.sub(
            rf"(?<!\[)\^{re.escape(fn_id)}(?!\])",
            f"[^{fn_id}]",
            text,
        )

    return text


def _extract_extra_images_from_zip(docx_path: Path, images_dir: Path, existing_count: int):
    try:
        with zipfile.ZipFile(docx_path, "r") as z:
            media_files = [f for f in z.namelist() if f.startswith("word/media/")]
            counter = existing_count
            for mf in media_files:
                data = z.read(mf)
                if len(data) == 0:
                    continue
                already = False
                for ef in images_dir.iterdir():
                    if ef.is_file() and ef.stat().st_size == len(data) and ef.read_bytes() == data:
                        already = True
                        break
                if not already:
                    ext = Path(mf).suffix.lower()
                    counter += 1
                    target = images_dir / f"img_{counter:03d}_extra{ext}"
                    with open(target, "wb") as f:
                        f.write(data)
    except Exception as e:
        print(f"  [WARN] Failed to extract extra images from docx: {e}")


def _convert_special_images(images_dir: Path) -> dict:
    """将 EMF/WMF/TIFF 等 AI 无法直接查看的格式转为 PNG。
    返回 {旧文件名: 新文件名} 映射，供调用方更新 Markdown 引用。"""
    converted = {}
    try:
        from PIL import Image
    except ImportError:
        return converted

    for img_file in list(images_dir.iterdir()):
        if img_file.suffix.lower() in CONVERTIBLE_EXTS:
            try:
                png_path = img_file.with_suffix(".png")
                if img_file.suffix.lower() in (".emf", ".wmf"):
                    _convert_emf_wmf_to_png(img_file, png_path)
                else:
                    img = Image.open(img_file)
                    img.save(png_path, "PNG")
                if png_path.exists() and png_path.stat().st_size > 0:
                    converted[img_file.name] = png_path.name
                    img_file.unlink()
            except Exception as e:
                print(f"  [WARN] Cannot convert {img_file.name} to PNG: {e}")
    return converted


def _convert_emf_wmf_to_png(src: Path, dst: Path):
    """尝试用 LibreOffice 或 Pillow 转换 EMF/WMF"""
    soffice = _find_libreoffice()
    if soffice:
        tmp_dir = tempfile.mkdtemp()
        try:
            cmd = [soffice, "--headless", "--convert-to", "png", "--outdir", tmp_dir, str(src)]
            subprocess.run(cmd, check=True, capture_output=True, timeout=30)
            converted = Path(tmp_dir) / (src.stem + ".png")
            if converted.exists():
                shutil.move(str(converted), str(dst))
                return
        except Exception:
            pass
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)

    try:
        from PIL import Image
        img = Image.open(src)
        img.save(dst, "PNG")
    except Exception:
        print(f"  [WARN] Cannot convert {src.name} to PNG. Install LibreOffice for EMF/WMF support.")


def _remove_empty_files(images_dir: Path):
    """删除 0 字节的垃圾文件"""
    for f in list(images_dir.iterdir()):
        if f.is_file() and f.stat().st_size == 0:
            f.unlink()


# ─── DOC 转换 ────────────────────────────────────────────────────────────────

_doc_convert_log: list[str] = []


def convert_doc(input_path: Path, output_dir: Path, embed_images: bool = False,
                progress_cb=None) -> Path:
    if progress_cb:
        progress_cb(5, "doc -> docx...")
    _doc_convert_log.clear()
    docx_path = _doc_to_docx(input_path, progress_cb)
    if docx_path is None:
        diag = "\n".join(_doc_convert_log) if _doc_convert_log else ""
        raise RuntimeError(
            ".doc 格式转换失败。\n"
            "\n"
            "【诊断信息】\n"
            f"{diag}\n"
            "\n"
            "【解决方案】请确保安装了以下任一软件：\n"
            "  1. Microsoft Office (推荐，需完整安装 Word 组件)\n"
            "  2. WPS Office: https://www.wps.cn/\n"
            "  3. LibreOffice: https://www.libreoffice.org/download/download/\n"
            "\n"
            "【常见问题】\n"
            "  - Office 365 在线版不支持 COM，需安装桌面版\n"
            "  - 安装后如仍失败，尝试以管理员身份运行\n"
            "  - 32位 Python + 64位 Office 可能不兼容，建议统一位数"
        )
    try:
        return convert_docx(docx_path, output_dir, embed_images, progress_cb)
    finally:
        if docx_path.exists():
            docx_path.unlink()


def _doc_to_docx(doc_path: Path, progress_cb=None) -> Path | None:
    result = _doc_to_docx_via_com(doc_path, progress_cb)
    if result:
        return result
    result = _doc_to_docx_via_doc2docx(doc_path, progress_cb)
    if result:
        return result
    return _doc_to_docx_via_libreoffice(doc_path, progress_cb)


def _doc_to_docx_via_com(doc_path: Path, progress_cb=None) -> Path | None:
    """方案1：pywin32 COM 接口调用 Word/WPS（Windows 专用）"""
    try:
        import win32com.client
    except ImportError:
        _doc_convert_log.append("[COM] pywin32 未安装，跳过 COM 方案")
        return None

    if progress_cb:
        progress_cb(8, "Word/WPS COM...")

    import platform
    py_bits = platform.architecture()[0]
    _doc_convert_log.append(f"[COM] Python {sys.version.split()[0]} ({py_bits})")

    _kill_office_processes()

    tmp_copy_dir = tempfile.mkdtemp()
    safe_name = f"input_doc{doc_path.suffix}"
    safe_path = Path(tmp_copy_dir) / safe_name
    shutil.copy2(str(doc_path), str(safe_path))
    abs_doc = str(safe_path.resolve()).replace("/", "\\")
    _doc_convert_log.append(f"[COM] 源文件: {doc_path.name} -> {abs_doc}")

    prog_ids = ["Word.Application", "Kwps.Application", "KWPS.Application"]
    last_err = None

    for prog_id in prog_ids:
        word = None
        doc = None
        try:
            word = win32com.client.Dispatch(prog_id)
            word.Visible = False
            word.DisplayAlerts = 0

            app_name = "Unknown"
            try:
                app_name = f"{word.Name} {word.Version}"
            except Exception:
                pass
            _doc_convert_log.append(f"[COM] {prog_id} -> {app_name}")

            tmp_dir = tempfile.mkdtemp()
            docx_out = Path(tmp_dir) / (doc_path.stem + ".docx")
            abs_out = str(docx_out.resolve()).replace("/", "\\")

            doc = word.Documents.Open(
                abs_doc,
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
                Visible=False,
            )

            try:
                doc.SaveAs2(abs_out, FileFormat=16)
            except AttributeError:
                doc.SaveAs(abs_out, FileFormat=16)

            doc.Close(False)
            doc = None

            if docx_out.exists() and docx_out.stat().st_size > 0:
                print(f"[OK] .doc -> .docx via COM ({prog_id})")
                _doc_convert_log.append(f"[COM] 成功: {prog_id}")
                return docx_out

        except Exception as e:
            last_err = e
            err_msg = str(e).replace("\n", " ")[:150]
            _doc_convert_log.append(f"[COM] {prog_id} 失败: {err_msg}")
            if doc:
                try:
                    doc.Close(False)
                except Exception:
                    pass
        finally:
            if word:
                try:
                    word.Quit()
                except Exception:
                    pass
            import time
            time.sleep(0.5)

    if last_err:
        print(f"[DOC COM] all COM ProgIDs failed, last: {last_err}")
    return None


def _kill_office_processes():
    """杀掉残留的 Word/WPS 后台进程，避免 COM 端口冲突和文件锁"""
    if sys.platform != "win32":
        return
    for proc_name in ("WINWORD", "wps"):
        try:
            subprocess.run(
                ["taskkill", "/F", "/IM", f"{proc_name}.exe"],
                capture_output=True, timeout=10,
            )
        except Exception:
            pass
    import time
    time.sleep(1)


def _doc_to_docx_via_doc2docx(doc_path: Path, progress_cb=None) -> Path | None:
    """方案2：doc2docx 库（封装了 COM，兼容性更好）"""
    try:
        from doc2docx import convert
    except ImportError:
        _doc_convert_log.append("[doc2docx] 未安装，跳过 (pip install doc2docx)")
        return None

    if progress_cb:
        progress_cb(8, "doc2docx...")

    tmp_dir = tempfile.mkdtemp()
    docx_out = Path(tmp_dir) / (doc_path.stem + ".docx")
    try:
        convert(str(doc_path.resolve()), str(docx_out.resolve()))
        if docx_out.exists() and docx_out.stat().st_size > 0:
            print("[OK] .doc -> .docx via doc2docx")
            _doc_convert_log.append("[doc2docx] 成功")
            return docx_out
    except Exception as e:
        err_msg = str(e).replace("\n", " ")[:150]
        _doc_convert_log.append(f"[doc2docx] 失败: {err_msg}")
        print(f"[DOC doc2docx] failed: {e}")
    return None


def _doc_to_docx_via_libreoffice(doc_path: Path, progress_cb=None) -> Path | None:
    """方案3：LibreOffice 命令行（跨平台备选）"""
    soffice = _find_libreoffice()
    if not soffice:
        _doc_convert_log.append("[LibreOffice] 未安装，跳过")
        return None

    if progress_cb:
        progress_cb(8, "LibreOffice...")

    _doc_convert_log.append(f"[LibreOffice] 找到: {soffice}")
    tmp_dir = tempfile.mkdtemp()
    try:
        env = os.environ.copy()
        env["HOME"] = tmp_dir
        env["USERPROFILE"] = tmp_dir
        cmd = [soffice, "--headless", "--norestore", "--convert-to", "docx",
               "--outdir", tmp_dir, str(doc_path)]
        proc = subprocess.run(cmd, capture_output=True, timeout=180, env=env)
        result_path = Path(tmp_dir) / (doc_path.stem + ".docx")
        if result_path.exists() and result_path.stat().st_size > 0:
            print("[OK] .doc -> .docx via LibreOffice")
            _doc_convert_log.append("[LibreOffice] 成功")
            return result_path
        if proc.returncode != 0:
            stderr = proc.stderr.decode("utf-8", errors="replace")[:200]
            _doc_convert_log.append(f"[LibreOffice] exit {proc.returncode}: {stderr}")
    except subprocess.TimeoutExpired:
        _doc_convert_log.append("[LibreOffice] 超时 (180s)")
    except (subprocess.CalledProcessError, FileNotFoundError, OSError) as e:
        _doc_convert_log.append(f"[LibreOffice] 异常: {e}")
    return None


def _find_libreoffice() -> str | None:
    if shutil.which("soffice"):
        return "soffice"
    for p in [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/usr/bin/soffice", "/usr/local/bin/soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]:
        if os.path.isfile(p):
            return p
    return None


# ─── PDF 转换 ────────────────────────────────────────────────────────────────

def convert_pdf(input_path: Path, output_dir: Path, embed_images: bool = False,
                progress_cb=None, engine: str = "auto") -> Path:
    stem = input_path.stem
    images_dir = output_dir / f"{stem}_images"
    images_dir.mkdir(parents=True, exist_ok=True)

    engine = engine or "auto"
    if engine == "auto":
        engine = _detect_best_pdf_engine()
        if progress_cb:
            progress_cb(5, f"engine: {engine}")

    fallback_order = ["mineru", "docling", "pymupdf4llm"]
    converters = {
        "mineru": _convert_pdf_mineru,
        "docling": _convert_pdf_docling,
        "pymupdf4llm": _convert_pdf_pymupdf4llm,
    }

    if engine in converters:
        start_idx = fallback_order.index(engine) if engine in fallback_order else 0
        engines_to_try = fallback_order[start_idx:]
    else:
        engines_to_try = [engine]

    markdown_text = None
    for eng in engines_to_try:
        converter_fn = converters.get(eng)
        if not converter_fn:
            continue
        try:
            if progress_cb:
                progress_cb(8, f"trying {eng}...")
            markdown_text = converter_fn(input_path, images_dir, progress_cb)
            engine = eng
            break
        except Exception as e:
            print(f"  [WARN] {eng} failed: {e}, trying next engine...")
            if progress_cb:
                progress_cb(8, f"{eng} failed, fallback...")
            continue

    if markdown_text is None:
        raise RuntimeError("All PDF engines failed. Check dependencies.")

    if progress_cb:
        progress_cb(70, "images...")

    _remove_empty_files(images_dir)
    image_count = len(list(images_dir.glob("*")))

    if embed_images:
        markdown_text = _embed_images_in_md(markdown_text, images_dir)

    header = _build_header(input_path, f"PDF ({engine})", image_count)
    markdown_text = header + markdown_text

    if progress_cb:
        progress_cb(90, "saving...")

    md_path = output_dir / f"{stem}.md"
    md_path.write_text(markdown_text, encoding="utf-8")

    print(f"[OK] PDF ({engine}) -> {md_path}")
    print(f"  images: {image_count}, dir: {images_dir}")

    if progress_cb:
        progress_cb(100, "done")
    return md_path


def _detect_best_pdf_engine() -> str:
    """自动检测可用的最佳 PDF 引擎，按准确率降级：MinerU > Docling > pymupdf4llm"""
    try:
        from mineru.cli.common import do_parse  # noqa: F401
        return "mineru"
    except ImportError:
        pass
    try:
        from docling.document_converter import DocumentConverter  # noqa: F401
        return "docling"
    except ImportError:
        pass
    try:
        import pymupdf4llm  # noqa: F401
        return "pymupdf4llm"
    except ImportError:
        pass
    raise RuntimeError(
        "No PDF engine available. Install one of:\n"
        "  pip install MinerU          # Best for Chinese docs (recommended)\n"
        "  pip install docling         # Best overall accuracy\n"
        "  pip install pymupdf4llm     # Fastest, basic accuracy"
    )


def _convert_pdf_mineru(input_path: Path, images_dir: Path, progress_cb=None) -> str:
    """使用 MinerU 引擎转换 PDF（中文文档最佳，表格准确率 0.87）"""
    try:
        from mineru.cli.common import do_parse
    except ImportError as e:
        raise RuntimeError("MinerU not installed: pip install MinerU") from e

    if progress_cb:
        progress_cb(10, "MinerU: parsing PDF...")

    output_dir = images_dir.parent / f"{input_path.stem}_mineru_tmp"
    output_dir.mkdir(parents=True, exist_ok=True)

    pdf_bytes = input_path.read_bytes()

    if progress_cb:
        progress_cb(20, "MinerU: model inference + extraction...")

    do_parse(
        output_dir=str(output_dir),
        pdf_file_names=[input_path.name],
        pdf_bytes_list=[pdf_bytes],
        p_lang_list=["ch"],
        backend="pipeline",
        parse_method="auto",
        table_enable=True,
        formula_enable=True,
        f_dump_md=True,
        f_dump_middle_json=False,
        f_dump_model_output=False,
        f_dump_orig_pdf=False,
        f_dump_content_list=False,
        f_draw_layout_bbox=False,
        f_draw_span_bbox=False,
    )

    if progress_cb:
        progress_cb(55, "MinerU: reading output...")

    md_files = list(output_dir.rglob("*.md"))
    if md_files:
        markdown_text = md_files[0].read_text(encoding="utf-8")
    else:
        raise RuntimeError("MinerU produced no markdown output")

    for img_file in output_dir.rglob("*"):
        if img_file.is_file() and img_file.suffix.lower() in AI_VIEWABLE_EXTS:
            shutil.copy2(str(img_file), str(images_dir / img_file.name))

    if progress_cb:
        progress_cb(65, "MinerU: cleanup...")

    shutil.rmtree(str(output_dir), ignore_errors=True)
    markdown_text = _clean_markdown(markdown_text)
    return markdown_text


def _convert_pdf_docling(input_path: Path, images_dir: Path, progress_cb=None) -> str:
    """使用 Docling 引擎转换 PDF（综合准确率最高 0.86，表格 0.89）"""
    try:
        from docling.document_converter import DocumentConverter
    except ImportError as e:
        raise RuntimeError("Docling not installed: pip install docling") from e

    if progress_cb:
        progress_cb(10, "Docling: loading models...")

    converter = DocumentConverter()

    if progress_cb:
        progress_cb(30, "Docling: converting...")

    result = converter.convert(str(input_path))

    if progress_cb:
        progress_cb(50, "Docling: exporting markdown...")

    markdown_text = result.document.export_to_markdown(image_mode="referenced")

    if progress_cb:
        progress_cb(60, "Docling: saving images...")

    for element, _level in result.document.iterate_items():
        if hasattr(element, "image") and element.image:
            try:
                pil_img = element.image.pil_image
                if pil_img:
                    img_name = f"{input_path.stem}-{element.self_ref}.png"
                    img_name = re.sub(r'[^\w\-.]', '_', img_name)
                    pil_img.save(str(images_dir / img_name))
            except Exception:
                pass

    if not any(images_dir.iterdir()):
        try:
            import fitz
            doc = fitz.open(str(input_path))
            img_idx = 0
            for page_idx, page in enumerate(doc):
                for img_info in page.get_images(full=True):
                    xref = img_info[0]
                    base_image = doc.extract_image(xref)
                    if base_image:
                        img_bytes = base_image["image"]
                        ext = "." + base_image.get("ext", "png")
                        img_name = f"{input_path.stem}-{page_idx}-{img_idx}{ext}"
                        (images_dir / img_name).write_bytes(img_bytes)
                        img_idx += 1
            doc.close()
        except Exception:
            pass

    markdown_text = _docling_post_process(markdown_text)
    return markdown_text


def _docling_post_process(text: str) -> str:
    """Docling 引擎专用后处理：合并拆分的代码块、修复格式异常"""
    text = _merge_split_code_blocks(text)
    text = _clean_markdown(text)
    return text


def _merge_split_code_blocks(text: str) -> str:
    """合并被拆分的相邻代码块（两个代码块之间仅有空行时合并为一个）"""
    text = re.sub(r"```\s*\n\s*\n\s*```", "", text)
    text = re.sub(r"```\s*\n```", "", text)
    return text


def _convert_pdf_pymupdf4llm(input_path: Path, images_dir: Path, progress_cb=None) -> str:
    """使用 pymupdf4llm 引擎转换 PDF（最快但准确率最低，含 fitz 补全）"""
    try:
        import pymupdf4llm
        import fitz  # noqa: F401
    except ImportError as e:
        raise RuntimeError("Missing dependencies: pip install pymupdf pymupdf4llm") from e

    if progress_cb:
        progress_cb(10, "fitz: text extraction...")

    raw_texts = _extract_pdf_all_texts(str(input_path))

    if progress_cb:
        progress_cb(30, "pymupdf4llm: converting...")

    markdown_text = pymupdf4llm.to_markdown(
        str(input_path), write_images=True,
        image_path=str(images_dir), image_format="png", dpi=200,
        table_strategy="lines",
        force_text=True,
    )

    if progress_cb:
        progress_cb(55, "post-processing...")

    markdown_text = _pdf_post_process(markdown_text, raw_texts)
    return markdown_text


# ─── PDF 后处理 ───────────────────────────────────────────────────────────────

def _extract_pdf_all_texts(pdf_path: str) -> list:
    """用 PyMuPDF (fitz) 逐页提取 PDF 中所有文本块，按阅读顺序排列。
    pymupdf4llm 会丢失文本框/灰色背景块中的内容，这里做全量提取作为补充数据源。"""
    import fitz
    all_texts = []
    doc = fitz.open(pdf_path)
    for page in doc:
        blocks = page.get_text("dict", sort=True)["blocks"]
        for block in blocks:
            if block["type"] != 0:
                continue
            for line in block.get("lines", []):
                spans_text = "".join(span["text"] for span in line.get("spans", []))
                cleaned = spans_text.strip()
                if cleaned:
                    all_texts.append(cleaned)
    doc.close()
    return all_texts


def _pdf_post_process(text: str, raw_texts: list = None) -> str:
    text = _rebuild_pdf_code_blocks(text)
    text = _fix_pdf_tables(text)
    text = _fix_pdf_headings(text)
    text = _fix_pdf_chinese_spaces(text)
    text = _fix_pdf_empty_sections(text)
    if raw_texts:
        text = _fill_missing_section_content(text, raw_texts)
    text = _clean_markdown(text)
    return text


def _rebuild_pdf_code_blocks(text: str) -> str:
    """彻底重建代码块：先剥离所有 ``` 围栏，再根据内容特征重新包裹。
    pymupdf4llm 对 PDF 的代码块围栏经常错乱（未闭合、吞标题、跨页断裂），
    与其修补不如重建。"""
    lines = text.split("\n")
    stripped_lines = []
    for line in lines:
        s = line.strip()
        if s == "```" or s == "``` " or re.match(r"^```\w*\s*$", s):
            stripped_lines.append("")
        else:
            stripped_lines.append(line)

    result = []
    in_code = False
    i = 0
    while i < len(stripped_lines):
        line = stripped_lines[i]
        stripped = line.strip()

        if in_code:
            if _is_structural_line(stripped):
                result.append("```")
                result.append("")
                in_code = False
                result.append(line)
                i += 1
                continue

            if stripped == "":
                j = i + 1
                while j < len(stripped_lines) and stripped_lines[j].strip() == "":
                    j += 1
                if j >= len(stripped_lines):
                    result.append("```")
                    in_code = False
                    for k in range(i, j):
                        result.append(stripped_lines[k])
                    i = j
                    continue
                next_s = stripped_lines[j].strip()
                if _is_structural_line(next_s) or not _looks_like_code(next_s, j, stripped_lines):
                    blank_count = j - i
                    if blank_count >= 2:
                        result.append("```")
                        in_code = False
                        for k in range(i, j):
                            result.append(stripped_lines[k])
                        i = j
                        continue
                result.append(line)
                i += 1
                continue

            result.append(line)
            i += 1
            continue

        if _is_structural_line(stripped) or stripped == "":
            result.append(line)
            i += 1
            continue

        if _looks_like_code(stripped, i, stripped_lines):
            result.append("")
            result.append("```")
            in_code = True
            result.append(line)
            i += 1
            continue

        result.append(line)
        i += 1

    if in_code:
        result.append("```")

    return "\n".join(result)


def _is_structural_line(stripped: str) -> bool:
    """判断是否为 Markdown 结构行（标题、表格、图片、加粗标签等），不应在代码块内"""
    if stripped.startswith("#"):
        return True
    if stripped.startswith("|") and stripped.endswith("|"):
        return True
    if stripped.startswith("!["):
        return True
    if re.match(r"^\*\*(.+?)\*\*\s*$", stripped):
        return True
    if stripped.startswith("注：") or stripped.startswith("注:"):
        return True
    if stripped == "解密后":
        return True
    if stripped.startswith("相似图像数据"):
        return True
    return False


def _looks_like_code(stripped: str, idx: int, lines: list) -> bool:
    """判断当前行是否像代码块的开始（JSON/Java 等代码特征）"""
    if stripped in ("{", "}", "{}", "});", "},"):
        return True
    if re.match(r"^\s*[{}\[\]]", stripped):
        return True
    if re.match(r'^\s*"[^"]+"\s*:', stripped):
        return True
    if re.match(r"^\s*(import|public|private|protected|class|static|try|catch|return|if|for|while)\s", stripped):
        return True
    if re.match(r"^\s*//", stripped):
        return True
    if re.match(r"^\s*(byte|String|Cipher|KeyGenerator|SecureRandom|SecretKeySpec)\s", stripped):
        return True
    if re.match(r"^\s*\w+\.\w+\(", stripped):
        return True
    if re.match(r"^\s*\w+\s*=\s*", stripped) and ";" in stripped:
        return True
    return False


def _fix_pdf_tables(text: str) -> str:
    """修复 PDF 提取表格的常见问题：
    1. 单元格内的 <br> 换行导致中文名称被拆行
    2. 列数不一致
    3. 表头行全空"""
    lines = text.split("\n")
    result = []
    table_lines = []
    in_table = False

    def flush_table():
        if not table_lines:
            return
        processed = []
        for tl in table_lines:
            cleaned = re.sub(r"<br\s*/?>", "", tl)
            processed.append(cleaned)

        col_counts = []
        for tl in processed:
            cols = tl.strip().split("|")
            if cols and cols[0].strip() == "":
                cols = cols[1:]
            if cols and cols[-1].strip() == "":
                cols = cols[:-1]
            col_counts.append(len(cols))

        if not col_counts:
            result.extend(processed)
            return

        max_cols = max(col_counts)
        for tl in processed:
            cols = tl.strip().split("|")
            if cols and cols[0].strip() == "":
                cols = cols[1:]
            if cols and cols[-1].strip() == "":
                cols = cols[:-1]
            while len(cols) < max_cols:
                cols.append(" ")
            is_sep = all(re.match(r"^\s*-{2,}\s*$", c) for c in cols)
            if is_sep:
                result.append("| " + " | ".join("---" for _ in cols) + " |")
            else:
                result.append("| " + " | ".join(c.strip() for c in cols) + " |")
        table_lines.clear()

    for line in lines:
        stripped = line.strip()
        if stripped.startswith("|") and stripped.endswith("|"):
            in_table = True
            table_lines.append(line)
        else:
            if in_table:
                flush_table()
                in_table = False
            result.append(line)

    if table_lines:
        flush_table()

    text = "\n".join(result)
    text = _fix_table_headers(text)
    return text


def _fix_pdf_headings(text: str) -> str:
    """修复 PDF 提取的标题格式：
    1. 去除标题中的 ** 加粗标记（# **标题** -> # 标题）
    2. 确保标题前后有空行
    3. 修正标题层级：基础约定/业务接口 -> ##，接口编号 -> ###，子节 -> ####"""
    lines = text.split("\n")
    result = []

    for i, line in enumerate(lines):
        stripped = line.strip()
        m = re.match(r"^(#{1,6})\s+\*\*(.+?)\*\*\s*$", stripped)
        if m:
            title = m.group(2).strip()
            level = _infer_heading_level(title)
            if result and result[-1].strip() != "":
                result.append("")
            result.append(f"{'#' * level} {title}")
            continue

        m2 = re.match(r"^(#{1,6})\s+(.+)$", stripped)
        if m2:
            title = m2.group(2).strip()
            level = _infer_heading_level(title)
            if result and result[-1].strip() != "":
                result.append("")
            result.append(f"{'#' * level} {title}")
            continue

        result.append(line)

    return "\n".join(result)


def _infer_heading_level(title: str) -> int:
    """根据标题内容推断正确的层级"""
    top_sections = ["基础约定", "业务接口"]
    for s in top_sections:
        if s in title:
            return 2

    if re.match(r"^\d+[\.\s]", title):
        return 3

    sub_sections = [
        "服务提供方", "接口地址", "请求参数", "响应参数",
        "请求方式", "Content-type", "加解密算法",
        "请求报文", "响应报文", "响应结果封装对象",
        "响应报文示例",
    ]
    for s in sub_sections:
        if s in title:
            return 4

    return 4


def _fix_pdf_chinese_spaces(text: str) -> str:
    """修复 PDF 提取时中文字符间被插入的多余空格。
    仅跳过代码块内的内容，表格和普通文本都处理。"""
    lines = text.split("\n")
    result = []
    in_code = False

    for line in lines:
        stripped = line.strip()
        if stripped.startswith("```"):
            in_code = not in_code
            result.append(line)
            continue

        if in_code:
            result.append(line)
            continue

        line = re.sub(r'([\u4e00-\u9fff])\s+([\u4e00-\u9fff])', r'\1\2', line)
        line = re.sub(r'([\u4e00-\u9fff])\s+([\u4e00-\u9fff])', r'\1\2', line)

        result.append(line)

    return "\n".join(result)



def _fix_pdf_empty_sections(text: str) -> str:
    """修复 PDF 提取中的孤立加粗标签。
    将 **服务提供方** 这类独立行转换为 #### 标题格式。
    同时将 **dataList属性** 这类转为加粗文本。"""
    section_keywords = [
        "服务提供方", "接口地址", "请求参数", "响应参数",
        "请求方式", "Content-type", "加解密算法",
        "请求报文", "响应报文", "响应结果封装对象",
        "响应报文示例",
    ]

    lines = text.split("\n")
    result = []

    for line in lines:
        stripped = line.strip()
        m = re.match(r"^\*\*(.+?)\*\*\s*$", stripped)
        if m:
            content = m.group(1).strip()
            is_section = False
            for kw in section_keywords:
                if content == kw or content.startswith(kw):
                    is_section = True
                    break
            if is_section:
                if result and result[-1].strip() != "":
                    result.append("")
                result.append(f"#### {content}")
                continue

        result.append(line)

    return "\n".join(result)


def _fill_missing_section_content(text: str, raw_texts: list) -> str:
    """用 fitz 提取的原始文本填充 pymupdf4llm 丢失的小节内容。
    1. 检测 #### 标题后面为空（紧跟另一个标题或空行）的情况，填入值
    2. 检测缺失的标题（如"接口地址"后直接跟"请求参数"，缺少"请求方式"），自动插入"""
    fillable_sections = {
        "服务提供方": True,
        "接口地址": True,
        "请求方式": True,
        "响应参数": True,
    }

    expected_order = ["服务提供方", "接口地址", "请求方式"]

    raw_lookup = _build_raw_text_lookup(raw_texts, fillable_sections)

    lines = text.split("\n")
    result = []
    i = 0
    interface_idx = 0
    seen_sections = set()

    while i < len(lines):
        stripped = lines[i].strip()

        m_iface = re.match(r"^###\s+(\d+)[\.\s]", stripped)
        if m_iface:
            interface_idx = int(m_iface.group(1))
            seen_sections = set()

        m_section = re.match(r"^####\s+(.+)$", stripped)
        if m_section and interface_idx > 0:
            section_name = m_section.group(1).strip()

            if section_name == "请求参数" and interface_idx > 0:
                for exp in expected_order:
                    if exp not in seen_sections:
                        value = raw_lookup.get((interface_idx, exp))
                        if value:
                            result.append(f"#### {exp}")
                            result.append("")
                            result.append("```")
                            result.append(value)
                            result.append("```")
                            result.append("")
                            seen_sections.add(exp)

            matched_key = None
            for key in fillable_sections:
                if section_name == key or section_name.startswith(key):
                    matched_key = key
                    break

            if matched_key:
                seen_sections.add(matched_key)

            if matched_key and fillable_sections[matched_key]:
                j = i + 1
                while j < len(lines) and lines[j].strip() == "":
                    j += 1
                next_is_empty = (
                    j >= len(lines) or
                    re.match(r"^#{1,4}\s+", lines[j].strip()) or
                    lines[j].strip().startswith("|") or
                    lines[j].strip().startswith("---")
                )

                if next_is_empty:
                    value = raw_lookup.get((interface_idx, matched_key))
                    if value:
                        result.append(lines[i])
                        result.append("")
                        result.append("```")
                        result.append(value)
                        result.append("```")
                        i += 1
                        continue

        result.append(lines[i])
        i += 1

    return "\n".join(result)


def _build_raw_text_lookup(raw_texts: list, section_keys: dict) -> dict:
    """从 fitz 提取的原始文本序列中，构建 (接口编号, 小节名) -> 值 的查找表。
    原始文本中的模式通常是：
      '1. 装/移机新增工单列表查询接口'
      '服务提供方'
      '智慧综调[IDS]'
      '接口地址'
      '/api/rest/assistant/order/queryNewSheets'
    """
    lookup = {}
    current_interface = 0

    normalized = []
    for t in raw_texts:
        cleaned = re.sub(r'\s+', '', t) if re.search(r'[\u4e00-\u9fff]', t) else t.strip()
        normalized.append((cleaned, t.strip()))

    for idx, (norm, original) in enumerate(normalized):
        m = re.match(r'^(\d+)[\.\s、．]', norm)
        if m:
            current_interface = int(m.group(1))
            continue

        if current_interface == 0:
            continue

        for key in section_keys:
            key_norm = re.sub(r'\s+', '', key)
            if norm == key_norm or norm.startswith(key_norm):
                for j in range(idx + 1, min(idx + 5, len(normalized))):
                    candidate_norm, candidate_orig = normalized[j]
                    if not candidate_orig:
                        continue
                    is_next_key = False
                    for k2 in section_keys:
                        k2_norm = re.sub(r'\s+', '', k2)
                        if candidate_norm == k2_norm or candidate_norm.startswith(k2_norm):
                            is_next_key = True
                            break
                    if is_next_key:
                        break
                    m2 = re.match(r'^(\d+)[\.\s、．]', candidate_norm)
                    if m2:
                        break
                    if re.match(r'^#{1,4}\s', candidate_orig):
                        break
                    if candidate_orig in ("请求参数", "响应参数", "响应报文", "请求报文"):
                        break

                    lookup[(current_interface, key)] = candidate_orig
                    break
                break

    return lookup


# ─── 工具函数 ────────────────────────────────────────────────────────────────

def _guess_image_ext(content_type: str) -> str:
    return {
        "image/png": ".png", "image/jpeg": ".jpg", "image/gif": ".gif",
        "image/bmp": ".bmp", "image/tiff": ".tiff", "image/svg+xml": ".svg",
        "image/x-wmf": ".wmf", "image/x-emf": ".emf",
    }.get(content_type, ".png")


def _clean_markdown(text: str) -> str:
    text = re.sub(r"\n{3,}", "\n\n", text)
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = text.strip() + "\n"
    return text


def _build_header(input_path: Path, doc_type: str, image_count: int) -> str:
    return f"---\nsource: {input_path.name}\ntype: {doc_type}\nimages: {image_count}\n---\n\n"


def _embed_images_in_md(markdown_text: str, images_dir: Path) -> str:
    for img_file in images_dir.glob("*"):
        if img_file.suffix.lower() in AI_VIEWABLE_EXTS:
            b64 = base64.b64encode(img_file.read_bytes()).decode("utf-8")
            mime = _get_mime(img_file.suffix)
            data_uri = f"data:{mime};base64,{b64}"
            markdown_text = markdown_text.replace(str(img_file), data_uri)
            markdown_text = markdown_text.replace(img_file.name, data_uri)
    return markdown_text


def _get_mime(ext: str) -> str:
    return {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
            ".gif": "image/gif", ".bmp": "image/bmp", ".webp": "image/webp"
            }.get(ext.lower(), "image/png")


# ─── CLI 入口 ────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="doc/docx/pdf -> Markdown + images",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="  python doc2md.py file.docx\n  python doc2md.py file.pdf -o ./out\n  python doc2md.py ./docs/",
    )
    parser.add_argument("input", help="doc/docx/pdf file or directory")
    parser.add_argument("-o", "--output", help="output directory", default=None)
    parser.add_argument("--embed-images", action="store_true", help="embed images as base64")
    parser.add_argument("--engine", choices=PDF_ENGINES, default="auto",
                        help="PDF engine: auto, mineru, docling, pymupdf4llm (default: auto)")

    args = parser.parse_args()
    input_path = Path(args.input).resolve()

    if not input_path.exists():
        print(f"Error: not found -> {input_path}")
        sys.exit(1)

    if input_path.is_dir():
        files = list(input_path.glob("*.doc")) + list(input_path.glob("*.docx")) + list(input_path.glob("*.pdf"))
        if not files:
            print(f"Error: no doc/docx/pdf in -> {input_path}")
            sys.exit(1)
        print(f"Found {len(files)} files\n")
        for f in sorted(files):
            _convert_single(f, args.output, args.embed_images, args.engine)
            print()
    else:
        _convert_single(input_path, args.output, args.embed_images, args.engine)


def _convert_single(input_path: Path, output: str | None, embed_images: bool, engine: str = "auto"):
    output_dir = Path(output).resolve() if output else input_path.parent
    output_dir.mkdir(parents=True, exist_ok=True)

    converters = {".docx": convert_docx, ".doc": convert_doc, ".pdf": convert_pdf}
    converter = converters.get(input_path.suffix.lower())
    if not converter:
        print(f"Skip: {input_path.name}")
        return

    print(f"Converting: {input_path.name}")
    try:
        if converter == convert_pdf:
            converter(input_path, output_dir, embed_images, engine=engine)
        else:
            converter(input_path, output_dir, embed_images)
    except Exception as e:
        print(f"Failed: {input_path.name} -> {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
