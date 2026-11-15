import argparse
import json
import re
import sys
from pathlib import Path

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph


HEADING_STYLE = re.compile(r"Heading([1-9])")
CLAUSE_NUMBER = re.compile(r"^(?:п\.\s*)?(\d+(?:\.\d+)+)[.)]?(?=\s|$)")
APPENDIX_HEADING = re.compile(r"^\s*приложение\b(?:\s*№?\s*(\d+[а-яa-z]?)(?!\w))?", re.I)
PLAIN_APPENDIX_HEADING = re.compile(
    r"^(?i:приложение)(?:\s*№?\s*\d+[а-яА-Яa-zA-Z]?)?"
    r"(?:\s*[:.\-–—]\s*.+|\s+[А-ЯЁA-Z].+)?$",
)
APPENDIX_REFERENCE = re.compile(r"\bприложени(?:е|я|ю|ем|и)\b(?:\s*№?\s*(\d+[а-яa-z]?)(?!\w))?", re.I)
TO_CONTRACT = re.compile(r"\s+к\s+(?:настоящему\s+)?договору\b", re.I)
SAMPLE_CONTRACTS = ("company-template-v1.docx", "vendor-differences-v1.docx")


def row_cells(row):
    """Read each cell once when it spans several grid columns."""
    cells = row.cells
    values = []
    column = 0
    while column < len(cells):
        cell = cells[column]
        values.append(cell.text)
        column += cell.grid_span
    return values


def extract_blocks(path: Path):
    """Yield paragraphs and table rows in their original order."""
    if path.resolve().suffix.lower() == ".pdf":
        yield from pdf_blocks(path.resolve().as_posix(), pdf_pages(path))
        return
    source = path.resolve().as_posix()
    document = Document(path)
    headings = {}
    table_number = 0

    for position, item in enumerate(document.iter_inner_content(), start=1):
        location = {"source": source, "body_position": position}
        if isinstance(item, Paragraph):
            raw_text = item.text
            text = raw_text.strip()
            if not text:
                continue

            heading = HEADING_STYLE.fullmatch(item.style.style_id)
            plain_appendix = (
                PLAIN_APPENDIX_HEADING.fullmatch(text)
                and not TO_CONTRACT.match(text, len("Приложение"))
                and not text.endswith((".", "!", "?"))
            )
            if heading or plain_appendix:
                level = int(heading.group(1)) if heading else 1
                headings = {depth: title for depth, title in headings.items() if depth < level}
                headings[level] = text

            number = CLAUSE_NUMBER.match(text)
            kind = "heading" if heading or plain_appendix else "clause" if number else "paragraph"
            yield {
                **location,
                "kind": kind,
                "headings": list(headings.values()),
                "clause_number": number.group(1) if kind == "clause" else None,
                "text": raw_text,
            }
        elif isinstance(item, Table):
            table_number += 1
            for row_number, row in enumerate(item.rows, start=1):
                cells = row_cells(row)
                if any(cell.strip() for cell in cells):
                    yield {
                        **location,
                        "kind": "table_row",
                        "headings": list(headings.values()),
                        "table_number": table_number,
                        "row_number": row_number,
                        "cells": cells,
                        "text": "\t".join(cells),
                    }


def pdf_pages(path):
    import pdfplumber

    if hasattr(path, "seek"):
        path.seek(0)
    with pdfplumber.open(path) as pdf:
        if len(pdf.pages) > 200:
            raise ValueError("В PDF больше 200 страниц.")
        for page in pdf.pages:
            lines = page.extract_text_lines(return_chars=False)
            if not lines and page.images:
                raise ValueError(f"На странице {page.page_number} нет текстового слоя.")
            tables = page.find_tables()
            elements = [(table.bbox[1], table.bbox[0], "table", table.extract()) for table in tables]
            for line in lines:
                x = (line["x0"] + line["x1"]) / 2
                y = (line["top"] + line["bottom"]) / 2
                if any(x0 <= x <= x1 and top <= y <= bottom for x0, top, x1, bottom in
                       (table.bbox for table in tables)):
                    continue
                elements.append((line["top"], line["x0"], "line", line["text"]))
            yield page.page_number, [(kind, value) for _, _, kind, value in sorted(elements, key=lambda item: item[:2])]
            page.close()


PDF_SECTION = re.compile(r"^(?:раздел\s+\d+\b|\d+[.)]\s+[А-ЯЁA-Z])", re.I)
PDF_APPENDIX = re.compile(r"^приложение\s*№?\s*\d+[а-яa-z]?\s+к\s+(?:настоящему\s+)?договору\b", re.I)


def pdf_blocks(source, pages):
    position = table_number = 0
    headings = []
    pending = None
    for page, elements in pages:
        for kind, value in elements:
            if kind == "table":
                if pending:
                    yield pending
                    pending = None
                table_number += 1
                for row_number, row in enumerate(value, 1):
                    cells = [" ".join((cell or "").split()) for cell in row]
                    if any(cells):
                        position += 1
                        yield dict(source=source, body_position=position, page=page, kind="table_row",
                                   headings=headings, table_number=table_number, row_number=row_number,
                                   cells=cells, text="\t".join(cells))
                continue
            text = " ".join(value.split())
            if not text or re.fullmatch(r"(?:стр\.?\s*)?\d+(?:\s*/\s*\d+)?", text, re.I):
                continue
            number = CLAUSE_NUMBER.match(text)
            appendix = PDF_APPENDIX.match(text) or PLAIN_APPENDIX_HEADING.fullmatch(text)
            heading = appendix or PDF_SECTION.match(text)
            if not number and not heading and pending and pending["kind"] == "clause":
                pending["text"] += " " + text
                if page != pending["page"]:
                    pending["page_end"] = page
                continue
            if pending:
                yield pending
            if heading:
                headings = [text] if appendix else [*headings[:1], text] if headings and APPENDIX_HEADING.match(headings[0]) else [text]
            position += 1
            pending = dict(source=source, body_position=position, page=page, headings=headings,
                           kind="heading" if heading else "clause" if number else "paragraph",
                           clause_number=number.group(1) if number and not heading else None, text=text)
    if pending:
        yield pending


def appendix_headings(blocks):
    for block in blocks:
        if block["kind"] == "heading" and (match := APPENDIX_HEADING.match(block["text"])):
            yield (match.group(1) or "").casefold(), block["text"].strip()


def missing_appendices(template_blocks, carrier_blocks):
    """Check references in the uploaded contract against its appendix headings."""
    template_headings = list(appendix_headings(template_blocks))
    template_titles = dict(template_headings)
    carrier_titles = dict(appendix_headings(carrier_blocks))

    missing = {}
    for block in carrier_blocks:
        if block["kind"] == "heading":
            continue
        for match in APPENDIX_REFERENCE.finditer(block["text"]):
            number = (match.group(1) or "").casefold()
            if not number:
                # An unnamed reference identifies an appendix only when the template gives one candidate.
                if carrier_titles or len(template_headings) != 1 or not TO_CONTRACT.match(block["text"], match.end()):
                    continue
                number = next(iter(template_titles))
            if number in carrier_titles:
                continue
            finding = missing.setdefault(number, {
                "title": template_titles.get(number, f"Приложение № {number}"),
                "references": [],
            })
            reference = {"clause_number": block.get("clause_number"), "quote": block["text"].strip()}
            reference.update({key: block[key] for key in ("body_position", "table_number", "row_number",
                                                         "page", "page_end") if key in block})
            if reference not in finding["references"]:
                finding["references"].append(reference)
    return list(missing.values())


def main():
    parser = argparse.ArgumentParser(description="Выводит блоки DOCX в консоль в формате JSON Lines.")
    parser.add_argument("docx", nargs="*", type=Path, help="Пути к договорам. По умолчанию берёт два учебных DOCX рядом со скриптом.")
    args = parser.parse_args()

    paths = args.docx or [Path(__file__).parent / "examples" / name for name in SAMPLE_CONTRACTS]
    for path in paths:
        if path.suffix.lower() != ".docx" or not path.is_file():
            parser.error(f"Нужен существующий файл DOCX: {path}")

    sys.stdout.reconfigure(encoding="utf-8")
    for path in paths:
        for block in extract_blocks(path):
            print(json.dumps(block, ensure_ascii=False))


if __name__ == "__main__":
    main()
