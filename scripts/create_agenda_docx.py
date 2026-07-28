#!/usr/bin/env python3
"""Create a Word agenda document from structured JSON.

The script uses only the Python standard library. It validates the send-ahead
bullets so agenda pre-reads stay concise: no more than 10 bullets, with each
bullet containing 5-10 words.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape


WORD_LIMIT_RE = re.compile(r"[A-Za-z0-9]+(?:'[A-Za-z0-9]+)?")


EXAMPLE_DATA = {
    "title": "1:1 with Manager - 2026-06-11",
    "subtitle": "Prepared agenda",
    "send_ahead_bullets": [
        "Align on top priorities for this week",
        "Review prior commitments and next steps",
        "Discuss feedback and growth opportunities today",
    ],
    "context_reviewed": ["Prior 1:1 notes", "Granola meeting notes"],
    "sections": [
        {
            "heading": "1. Check-in",
            "items": [
                {
                    "label": "Energy/morale",
                    "body": "Capture current energy, capacity, and any context affecting focus.",
                },
                {
                    "label": "Anything notable personally/professionally",
                    "body": "Share relevant changes without turning the meeting into status.",
                },
            ],
        },
        {
            "heading": "2. Five words",
            "items": [
                {"label": "My five", "body": "Priorities, blocker, feedback, stakeholders, growth."},
                {"label": "Manager's five", "body": "Ask for their five and merge the agenda."},
            ],
        },
    ],
}


def text(value: object) -> str:
    return "" if value is None else str(value)


def xml_text(value: object) -> str:
    return escape(text(value), {'"': "&quot;"})


def count_words(value: str) -> int:
    return len(WORD_LIMIT_RE.findall(value))


def validate_bullets(bullets: list[str]) -> None:
    if len(bullets) > 10:
        raise ValueError(f"send_ahead_bullets has {len(bullets)} bullets; maximum is 10")
    errors = []
    for index, bullet in enumerate(bullets, start=1):
        words = count_words(bullet)
        if words < 5 or words > 10:
            errors.append(f"bullet {index} has {words} words: {bullet!r}")
    if errors:
        raise ValueError("Send-ahead bullets must be 5-10 words each:\n" + "\n".join(errors))


def run_xml(value: str, *, bold: bool = False, italic: bool = False, size: int | None = None) -> str:
    props = []
    if bold:
        props.append("<w:b/>")
    if italic:
        props.append("<w:i/>")
    if size:
        props.append(f'<w:sz w:val="{size}"/>')
    prop_xml = f"<w:rPr>{''.join(props)}</w:rPr>" if props else ""
    space = ' xml:space="preserve"' if value[:1].isspace() or value[-1:].isspace() else ""
    return f"<w:r>{prop_xml}<w:t{space}>{xml_text(value)}</w:t></w:r>"


def paragraph_xml(
    runs: list[str],
    *,
    style: str | None = None,
    bullet: bool = False,
    spacing_after: int = 120,
) -> str:
    p_props = [f'<w:spacing w:after="{spacing_after}"/>']
    if style:
        p_props.append(f'<w:pStyle w:val="{style}"/>')
    if bullet:
        p_props.append('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>')
    return f"<w:p><w:pPr>{''.join(p_props)}</w:pPr>{''.join(runs)}</w:p>"


def simple_paragraph(value: str, *, style: str | None = None, bullet: bool = False) -> str:
    return paragraph_xml([run_xml(value)], style=style, bullet=bullet)


def labeled_paragraph(label: str, body: str) -> str:
    runs = [run_xml(label.rstrip(":") + ": ", bold=True)]
    if body:
        runs.append(run_xml(body))
    return paragraph_xml(runs)


def normalize_items(raw_items: object) -> list[dict[str, str]]:
    if raw_items is None:
        return []
    if isinstance(raw_items, list):
        normalized = []
        for item in raw_items:
            if isinstance(item, dict):
                normalized.append({"label": text(item.get("label")), "body": text(item.get("body"))})
            else:
                normalized.append({"label": "", "body": text(item)})
        return normalized
    if isinstance(raw_items, dict):
        return [{"label": text(key), "body": text(value)} for key, value in raw_items.items()]
    return [{"label": "", "body": text(raw_items)}]


def document_body(data: dict) -> str:
    title = text(data.get("title") or "1:1 Agenda")
    subtitle = text(data.get("subtitle") or "")
    bullets = [text(item) for item in data.get("send_ahead_bullets", [])]
    validate_bullets(bullets)

    parts = [simple_paragraph(title, style="Title")]
    if subtitle:
        parts.append(simple_paragraph(subtitle, style="Subtitle"))

    if bullets:
        parts.append(simple_paragraph("Send-ahead bullets", style="Heading1"))
        for bullet in bullets:
            parts.append(simple_paragraph(bullet, bullet=True))

    context_reviewed = [text(item) for item in data.get("context_reviewed", []) if text(item).strip()]
    if context_reviewed:
        parts.append(simple_paragraph("Context reviewed", style="Heading1"))
        for item in context_reviewed:
            parts.append(simple_paragraph(item, bullet=True))

    for section in data.get("sections", []):
        if isinstance(section, dict):
            heading = text(section.get("heading") or section.get("title") or "Section")
            items = normalize_items(section.get("items"))
            notes = [text(item) for item in section.get("notes", [])]
        else:
            heading = text(section)
            items = []
            notes = []

        parts.append(simple_paragraph(heading, style="Heading1"))
        for item in items:
            label = item.get("label", "")
            body = item.get("body", "")
            if label:
                parts.append(labeled_paragraph(label, body))
            elif body:
                parts.append(simple_paragraph(body))
        for note in notes:
            if note.strip():
                parts.append(simple_paragraph(note, bullet=True))

    footer_notes = [text(item) for item in data.get("notes", []) if text(item).strip()]
    if footer_notes:
        parts.append(simple_paragraph("Additional notes", style="Heading1"))
        for note in footer_notes:
            parts.append(simple_paragraph(note, bullet=True))

    return "\n".join(parts)


def content_types_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>
"""


def root_rels_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""


def document_rels_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>
"""


def styles_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
    <w:rPr><w:sz w:val="22"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:rPr><w:b/><w:sz w:val="34"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Subtitle">
    <w:name w:val="Subtitle"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:rPr><w:i/><w:sz w:val="24"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr><w:spacing w:before="240" w:after="120"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="28"/></w:rPr>
  </w:style>
</w:styles>
"""


def numbering_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:multiLevelType w:val="singleLevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="&#8226;"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>
"""


def document_xml(data: dict) -> str:
    body = document_body(data)
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    {body}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>
"""


def create_docx(data: dict, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as docx:
        docx.writestr("[Content_Types].xml", content_types_xml())
        docx.writestr("_rels/.rels", root_rels_xml())
        docx.writestr("word/_rels/document.xml.rels", document_rels_xml())
        docx.writestr("word/document.xml", document_xml(data))
        docx.writestr("word/styles.xml", styles_xml())
        docx.writestr("word/numbering.xml", numbering_xml())


def load_data(args: argparse.Namespace) -> dict:
    if args.example:
        return EXAMPLE_DATA
    if not args.input:
        raise ValueError("Provide --input or --example")
    # utf-8-sig tolerates a BOM, which PowerShell's `Set-Content -Encoding utf8`
    # (Windows PowerShell 5.1) prepends; plain utf-8 would reject it.
    with Path(args.input).open("r", encoding="utf-8-sig") as handle:
        data = json.load(handle)
    if not isinstance(data, dict):
        raise ValueError("Input JSON must be an object")
    return data


def main() -> int:
    parser = argparse.ArgumentParser(description="Create a 1:1 agenda Word document.")
    parser.add_argument("--input", help="Path to agenda JSON")
    parser.add_argument("--output", required=True, help="Path to write .docx")
    parser.add_argument("--example", action="store_true", help="Write an example agenda document")
    args = parser.parse_args()

    try:
        data = load_data(args)
        create_docx(data, Path(args.output))
    except Exception as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1

    print(f"Wrote {Path(args.output).resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
