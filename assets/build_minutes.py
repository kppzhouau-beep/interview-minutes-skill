#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Build a Word interview memo from the bundled enterprise template.

The script only replaces content. It never normalizes fonts or recreates styles.
Every output is checked by scripts/verify_docx_format.py before success is reported.
"""

from __future__ import annotations

import argparse
from datetime import datetime
import json
import re
import subprocess
import sys
import zipfile
from copy import deepcopy
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
DEFAULT_TEMPLATE = HERE / "template_minutes.docx"
VERIFIER = ROOT / "scripts" / "verify_docx_format.py"
CRITICAL_TEMPLATE_PARTS = {
    "word/styles.xml",
    "word/numbering.xml",
    "word/settings.xml",
    "word/theme/theme1.xml",
}
META_LABELS = (
    "访谈时间",
    "访谈地点",
    "访谈来源",
    "访谈对象及其背景介绍",
    "访谈人员",
)


def number_id(paragraph):
    node = paragraph._p.find(".//" + qn("w:numId"))
    return node.get(qn("w:val")) if node is not None else None


def clear_inline_content(paragraph_element):
    for child in list(paragraph_element):
        if child.tag != qn("w:pPr"):
            paragraph_element.remove(child)


def clone_run(run_element, text):
    new_run = deepcopy(run_element)
    for child in list(new_run):
        if child.tag != qn("w:rPr"):
            new_run.remove(child)
    text_node = OxmlElement("w:t")
    if text.startswith(" ") or text.endswith(" "):
        text_node.set(qn("xml:space"), "preserve")
    text_node.text = text
    new_run.append(text_node)
    return new_run


def clone_simple(paragraph, text, run_index=0):
    new_p = deepcopy(paragraph._p)
    clear_inline_content(new_p)
    new_p.append(clone_run(paragraph.runs[run_index]._r, text))
    return new_p


def clone_split(paragraph, lead, body):
    new_p = deepcopy(paragraph._p)
    clear_inline_content(new_p)
    new_p.append(clone_run(paragraph.runs[0]._r, lead))
    if body:
        body_template = paragraph.runs[1]._r if len(paragraph.runs) > 1 else paragraph.runs[0]._r
        new_p.append(clone_run(body_template, body))
    return new_p


def add_page_break_before(paragraph_element):
    ppr = paragraph_element.find(qn("w:pPr"))
    if ppr is None:
        ppr = OxmlElement("w:pPr")
        paragraph_element.insert(0, ppr)
    if ppr.find(qn("w:pageBreakBefore")) is None:
        page_break = OxmlElement("w:pageBreakBefore")
        insertion_index = 0
        allowed_before = {qn("w:pStyle"), qn("w:keepNext"), qn("w:keepLines")}
        for index, child in enumerate(list(ppr)):
            if child.tag in allowed_before:
                insertion_index = index + 1
            else:
                break
        ppr.insert(insertion_index, page_break)


def clone_meta(paragraph, label, value, value_run):
    new_p = deepcopy(paragraph._p)
    clear_inline_content(new_p)
    new_p.append(clone_run(paragraph.runs[0]._r, label + "："))
    if value:
        new_p.append(clone_run(value_run._r, value))
    return new_p


def find_meta_prototype(paragraphs, label):
    return next((p for p in paragraphs if p.text.startswith(label)), None)


def normalize_item(item, *, split=False):
    if split:
        if isinstance(item, dict):
            return str(item.get("lead", "")), str(item.get("body", ""))
        if isinstance(item, (list, tuple)) and len(item) == 2:
            return str(item[0]), str(item[1])
        raise ValueError("结论和正文要点必须提供 lead 与 body")
    if isinstance(item, dict):
        return str(item.get("text", ""))
    return str(item)


def safe_component(value: str, field: str) -> str:
    value = value.strip()
    if not value:
        raise ValueError(f"{field}不能为空")
    if re.search(r"[\\/:*?\"<>|\x00-\x1f]", value):
        raise ValueError(f"{field}含文件名非法字符")
    return value


def infer_date_compact(payload: dict, meta: dict[str, str]) -> str:
    explicit = str(payload.get("date_compact", "")).strip()
    if re.fullmatch(r"\d{8}", explicit):
        datetime.strptime(explicit, "%Y%m%d")
        return explicit
    source = meta.get("访谈时间", "")
    match = re.search(r"(20\d{2})年\s*(\d{1,2})月\s*(\d{1,2})日", source)
    if not match:
        raise ValueError("必须提供 date_compact=YYYYMMDD，或在访谈时间中写明完整年月日")
    year, month, day = map(int, match.groups())
    return f"{year:04d}{month:02d}{day:02d}"


def restore_critical_template_parts(output: Path, reference: Path) -> None:
    """Undo serializer normalization of immutable template-level OOXML parts."""
    temporary = output.with_suffix(".restoring.docx")
    with zipfile.ZipFile(reference) as source_reference:
        preserved = {
            name: source_reference.read(name)
            for name in CRITICAL_TEMPLATE_PARTS
            if name in source_reference.namelist()
        }
    with zipfile.ZipFile(output) as source_output, zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as target:
        for info in source_output.infolist():
            data = preserved.get(info.filename, source_output.read(info.filename))
            target.writestr(info, data)
    temporary.replace(output)


def build(payload: dict, reference: Path, output_dir: Path) -> Path:
    project = safe_component(str(payload.get("project", "")), "project")
    subject = safe_component(str(payload.get("subject", "")), "subject")
    interview_type = safe_component(str(payload.get("interview_type", "专家访谈")), "interview_type")
    if not interview_type.endswith("访谈"):
        raise ValueError("interview_type必须以“访谈”结尾，例如“专家访谈”或“客户访谈”")
    meta_input = payload.get("meta") or {}
    meta = {label: str(meta_input.get(label, "")).strip() for label in META_LABELS}
    date_compact = infer_date_compact(payload, meta)
    title = f"{project}-{subject}{interview_type}"
    filename = f"{project}_{subject}{interview_type}纪要_{date_compact}.docx"

    conclusions = payload.get("conclusions") or []
    sections = payload.get("sections") or []
    if not conclusions:
        raise ValueError("conclusions不能为空")
    if not sections:
        raise ValueError("sections不能为空")
    kinds = [str(item.get("kind", "")) for item in sections if isinstance(item, dict)]
    for required in ("section", "subsection", "point"):
        if required not in kinds:
            raise ValueError(f"sections至少需要一个{required}")

    doc = Document(reference)
    paragraphs = doc.paragraphs
    body = doc.element.body
    table_element = next(child for child in body if child.tag == qn("w:tbl"))
    section_properties = body.find(qn("w:sectPr"))
    prototype = {
        "title": paragraphs[0],
        "h1_main": next(p for p in paragraphs if p.text.strip() == "主要结论"),
        "h1_minutes": next(p for p in paragraphs if p.text.strip() == "访谈纪要"),
        "empty": next((p for p in paragraphs if not p.text and number_id(p) is None), None),
        "conclusion": next(p for p in paragraphs if number_id(p) == "7"),
        "section": next(p for p in paragraphs if number_id(p) == "13"),
        "subsection": next(p for p in paragraphs if number_id(p) == "10"),
        "point": next(p for p in paragraphs if number_id(p) == "4"),
    }

    for child in list(body):
        if child.tag == qn("w:p"):
            body.remove(child)
    body.insert(body.index(table_element), clone_simple(prototype["title"], title))

    def append(element):
        if section_properties is not None:
            section_properties.addprevious(element)
        else:
            body.append(element)

    append(clone_simple(prototype["h1_main"], "主要结论"))
    for item in conclusions:
        lead, text = normalize_item(item, split=True)
        if not (lead + text).strip():
            raise ValueError("主要结论不得为空")
        append(clone_split(prototype["conclusion"], lead, text))

    if prototype["empty"] is not None:
        append(deepcopy(prototype["empty"]._p))

    minutes_heading = clone_simple(prototype["h1_minutes"], "访谈纪要")
    add_page_break_before(minutes_heading)
    append(minutes_heading)

    for item in sections:
        if not isinstance(item, dict):
            raise ValueError("sections中的每项必须是对象")
        kind = str(item.get("kind", ""))
        if kind in {"section", "subsection"}:
            text = normalize_item(item).strip()
            if not text:
                raise ValueError(f"{kind}文本不得为空")
            append(clone_simple(prototype[kind], text))
        elif kind == "point":
            lead, text = normalize_item(item, split=True)
            if not (lead + text).strip():
                raise ValueError("正文要点不得为空")
            append(clone_split(prototype["point"], lead, text))
        elif kind == "empty":
            append(deepcopy(prototype["empty"]._p) if prototype["empty"] is not None else OxmlElement("w:p"))
        else:
            raise ValueError(f"未知段落类型：{kind}")

    cell = doc.tables[0].rows[0].cells[0]
    old_meta = list(cell.paragraphs)
    fallback_value = next(
        (r for p in old_meta for r in p.runs if r.text.strip() and not r.bold),
        old_meta[0].runs[0],
    )
    new_meta = []
    for label in META_LABELS:
        paragraph = find_meta_prototype(old_meta, label) or old_meta[0]
        value_template = next((r for r in paragraph.runs[1:] if not r.bold), fallback_value)
        new_meta.append(clone_meta(paragraph, label, meta[label], value_template))
    for node in list(cell._tc.findall(qn("w:p"))):
        cell._tc.remove(node)
    for node in new_meta:
        cell._tc.append(node)

    properties = doc.core_properties
    properties.author = "Interview Minutes contributors"
    properties.last_modified_by = "Interview Minutes contributors"
    properties.title = title
    properties.subject = "Institutional interview minutes"
    properties.keywords = "interview, minutes, due diligence"
    properties.comments = ""

    output_dir.mkdir(parents=True, exist_ok=True)
    output = output_dir / filename
    doc.save(output)
    restore_critical_template_parts(output, reference)

    verify_cmd = [
        sys.executable,
        str(VERIFIER),
        str(output),
        "--reference",
        str(reference),
        "--project",
        project,
        "--subject",
        subject,
        "--interview-type",
        interview_type,
        "--date-compact",
        date_compact,
        "--expected-meta",
        json.dumps(meta, ensure_ascii=False),
    ]
    result = subprocess.run(verify_cmd, text=True, capture_output=True)
    if result.returncode != 0:
        sys.stderr.write(result.stdout)
        sys.stderr.write(result.stderr)
        raise RuntimeError("输出未通过强制格式校验，不得交付")
    print(result.stdout.strip())
    print(f"SAVED {output}")
    return output


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--payload", type=Path, required=True, help="UTF-8 JSON content payload")
    parser.add_argument("--reference", type=Path, default=DEFAULT_TEMPLATE)
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    payload = json.loads(args.payload.read_text(encoding="utf-8"))
    try:
        build(payload, args.reference, args.output_dir)
    except Exception as exc:
        print(f"BUILD FAILED: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
