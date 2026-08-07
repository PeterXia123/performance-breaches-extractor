#!/usr/bin/env python3

import argparse
import csv
import json
import re
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional
from xml.etree import ElementTree as ET


WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": WORD_NS}

W_P = "{%s}p" % WORD_NS
W_TBL = "{%s}tbl" % WORD_NS
W_R = "{%s}r" % WORD_NS
W_T = "{%s}t" % WORD_NS
W_TAB = "{%s}tab" % WORD_NS
W_BR = "{%s}br" % WORD_NS
W_CR = "{%s}cr" % WORD_NS
W_COLOR = "{%s}color" % WORD_NS
W_VAL = "{%s}val" % WORD_NS
W_THEME_COLOR = "{%s}themeColor" % WORD_NS

SEVERITY_RE = re.compile(
    r"^\s*(?P<number>\d+)\.\s*Severity of Breach Identified:\s*(?P<severity>.+?)\s*$",
    re.IGNORECASE,
)
GROUP_RE = re.compile(r"^[A-Za-z0-9&/\- ,()]+ Models$", re.IGNORECASE)
PLACEHOLDER_PREFIXES = (
    "select date",
    "click or tap to enter a date",
    "select remediation timeline",
    "select remediation date",
)
BLACK_THEME_COLORS = {"text1", "tx1", "dark1", "dk1"}
BLACK_COLOR_VALUES = {"000000", "000001", "auto"}
OUTPUT_COLUMNS = [
    "group_name",
    "chunk_number",
    "severity_raw",
    "record_status",
    "model_risk_rating",
    "model_ids",
    "model_names",
    "explanation_of_breach",
    "first_breach_identified",
    "second_breach_identified",
    "third_breach_identified",
    "plan_to_address",
    "remediation_date",
]


@dataclass
class RunInfo:
    text: str
    color_value: Optional[str]
    theme_color: Optional[str]


@dataclass
class ParagraphBlock:
    text: str
    runs: List[RunInfo]

    kind: str = "paragraph"


@dataclass
class TableBlock:
    rows: List[List[str]]

    kind: str = "table"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Extract Performance Breaches records from a DOCX and write "
            "CSV/JSON outputs without requiring package installation."
        )
    )
    parser.add_argument("input_file", help="Path to the source .docx file.")
    parser.add_argument(
        "--output-dir",
        default="out",
        help="Directory where breach_records.csv and breach_records.json will be written.",
    )
    parser.add_argument(
        "--csv-name",
        default="breach_records.csv",
        help="Output CSV filename inside the output directory.",
    )
    parser.add_argument(
        "--json-name",
        default="breach_records.json",
        help="Output JSON filename inside the output directory.",
    )
    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    input_path = Path(args.input_file)
    if input_path.suffix.lower() != ".docx":
        parser.error("This first version supports .docx input only.")

    records = extract_breach_records(input_path)

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    csv_path = output_dir / args.csv_name
    json_path = output_dir / args.json_name

    write_csv(records, csv_path)
    write_json(records, json_path)

    print("Wrote %s" % csv_path)
    print("Wrote %s" % json_path)
    print("Extracted %d breach record(s)" % len(records))


def extract_breach_records(docx_path: Path) -> List[Dict[str, str]]:
    blocks = parse_docx_blocks(docx_path)
    return build_records_from_blocks(blocks)


def parse_docx_blocks(docx_path: Path) -> List[object]:
    with zipfile.ZipFile(str(docx_path)) as archive:
        try:
            document_xml = archive.read("word/document.xml")
        except KeyError:
            raise RuntimeError("word/document.xml was not found in the DOCX.") from None

    root = ET.fromstring(document_xml)
    body = root.find("w:body", NS)
    if body is None:
        return []

    blocks: List[object] = []
    for child in list(body):
        if child.tag == W_P:
            blocks.append(parse_paragraph_block(child))
        elif child.tag == W_TBL:
            blocks.append(parse_table_block(child))
    return blocks


def parse_paragraph_block(paragraph_element: ET.Element) -> ParagraphBlock:
    runs: List[RunInfo] = []
    parts: List[str] = []
    for run_element in paragraph_element.iter(W_R):
        text = extract_run_text(run_element)
        if not text:
            continue
        color_value = None
        theme_color = None
        color_element = run_element.find("w:rPr/w:color", NS)
        if color_element is not None:
            color_value = normalize_case(color_element.get(W_VAL))
            theme_color = normalize_case(color_element.get(W_THEME_COLOR))
        runs.append(RunInfo(text=text, color_value=color_value, theme_color=theme_color))
        parts.append(text)

    text = clean_multiline_text("".join(parts))
    return ParagraphBlock(text=text, runs=runs)


def parse_table_block(table_element: ET.Element) -> TableBlock:
    rows: List[List[str]] = []
    for row_element in table_element.findall("./w:tr", NS):
        row: List[str] = []
        for cell_element in row_element.findall("./w:tc", NS):
            paragraphs = []
            for paragraph_element in cell_element.findall(".//w:p", NS):
                paragraph = parse_paragraph_block(paragraph_element)
                if paragraph.text:
                    paragraphs.append(paragraph.text)
            cell_text = "\n".join(paragraphs).strip()
            row.append(cell_text)
        if any(normalize_inline(cell) for cell in row):
            rows.append(row)
    return TableBlock(rows=rows)


def extract_run_text(run_element: ET.Element) -> str:
    parts: List[str] = []
    for child in list(run_element):
        if child.tag == W_T:
            parts.append(child.text or "")
        elif child.tag == W_TAB:
            parts.append("\t")
        elif child.tag in (W_BR, W_CR):
            parts.append("\n")
    return "".join(parts)


def build_records_from_blocks(blocks: Iterable[object]) -> List[Dict[str, str]]:
    records: List[Dict[str, str]] = []
    current_group = ""
    current_chunk: Optional[Dict[str, object]] = None
    in_breach_section = False

    for block in blocks:
        if isinstance(block, ParagraphBlock):
            text = normalize_inline(block.text)
            if not text:
                continue

            if not in_breach_section and "performance breaches:" in text.lower():
                in_breach_section = True
                continue

            if not in_breach_section:
                continue

            severity_match = SEVERITY_RE.match(text)
            if severity_match:
                if current_chunk is not None:
                    records.append(finalize_chunk(current_chunk))
                current_chunk = {
                    "group_name": current_group,
                    "chunk_number": severity_match.group("number"),
                    "severity_raw": normalize_severity_text(severity_match.group("severity")),
                    "record_status": classify_record_status(block),
                    "table": None,
                }
                continue

            if GROUP_RE.match(text):
                current_group = text
                continue

        elif isinstance(block, TableBlock):
            if in_breach_section and current_chunk is not None and current_chunk.get("table") is None:
                current_chunk["table"] = block.rows

    if current_chunk is not None:
        records.append(finalize_chunk(current_chunk))

    return records


def finalize_chunk(chunk: Dict[str, object]) -> Dict[str, str]:
    record = {column: "" for column in OUTPUT_COLUMNS}
    record["group_name"] = chunk.get("group_name", "") or ""
    record["chunk_number"] = str(chunk.get("chunk_number", "") or "")
    record["severity_raw"] = chunk.get("severity_raw", "") or ""
    record["record_status"] = chunk.get("record_status", "") or "decommissioned"

    table_rows = chunk.get("table") or []
    if table_rows:
        extracted = extract_fields_from_table(table_rows)
        record.update(extracted)

    return record


def extract_fields_from_table(rows: List[List[str]]) -> Dict[str, str]:
    extracted = {
        "model_risk_rating": "",
        "model_ids": "",
        "model_names": "",
        "explanation_of_breach": "",
        "first_breach_identified": "",
        "second_breach_identified": "",
        "third_breach_identified": "",
        "plan_to_address": "",
        "remediation_date": "",
    }

    for row in rows:
        row_lower = [normalize_inline(cell).lower() for cell in row]

        if any("model(s) affected" in cell for cell in row_lower):
            values = non_label_cells(row, "model(s) affected")
            if len(values) >= 1:
                extracted["model_risk_rating"] = normalize_scalar(values[0])
            if len(values) >= 2:
                extracted["model_ids"] = collapse_multivalue_cell(values[1])
            if len(values) >= 3:
                extracted["model_names"] = collapse_multivalue_cell(values[2])
            elif len(values) == 2 and not extracted["model_names"]:
                extracted["model_names"] = collapse_multivalue_cell(values[1])

        if any("explanation of breach" in cell for cell in row_lower):
            values = non_label_cells(row, "explanation of breach")
            extracted["explanation_of_breach"] = normalize_block_text(" ".join(values))

    date_header_index = find_row_index(rows, "1st breach identified")
    if date_header_index != -1 and date_header_index + 1 < len(rows):
        header_row = rows[date_header_index]
        value_row = rows[date_header_index + 1]
        extracted.update(extract_breach_dates(header_row, value_row))

    plan_row_index = find_plan_row_index(rows)
    if plan_row_index != -1:
        label_row = rows[plan_row_index]
        value_row = rows[plan_row_index + 1] if plan_row_index + 1 < len(rows) else []
        extracted.update(extract_plan_and_remediation(label_row, value_row))

    return extracted


def extract_breach_dates(header_row: List[str], value_row: List[str]) -> Dict[str, str]:
    extracted = {
        "first_breach_identified": "",
        "second_breach_identified": "",
        "third_breach_identified": "",
    }
    for index, header in enumerate(header_row):
        label = normalize_inline(header).lower()
        value = normalize_scalar(value_row[index] if index < len(value_row) else "")
        if "1st breach identified" in label:
            extracted["first_breach_identified"] = value
        elif "2nd breach identified" in label:
            extracted["second_breach_identified"] = value
        elif "3rd breach identified" in label:
            extracted["third_breach_identified"] = value
    return extracted


def extract_plan_and_remediation(label_row: List[str], value_row: List[str]) -> Dict[str, str]:
    extracted = {
        "plan_to_address": "",
        "remediation_date": "",
    }
    has_remediation_label = any(
        "remediation date" in normalize_inline(cell).lower() for cell in label_row
    )

    if value_row:
        if has_remediation_label and len(value_row) > 1:
            extracted["plan_to_address"] = normalize_block_text("\n".join(value_row[:-1]))
            extracted["remediation_date"] = normalize_scalar(value_row[-1])
        else:
            extracted["plan_to_address"] = normalize_block_text("\n".join(value_row))

    return extracted


def find_row_index(rows: List[List[str]], needle: str) -> int:
    needle_lower = needle.lower()
    for index, row in enumerate(rows):
        if any(needle_lower in normalize_inline(cell).lower() for cell in row):
            return index
    return -1


def find_plan_row_index(rows: List[List[str]]) -> int:
    targets = (
        "mva comments/plan to address persistent breach",
        "plan to address persistent breach",
    )
    for index, row in enumerate(rows):
        normalized = [normalize_inline(cell).lower() for cell in row]
        if any(any(target in cell for target in targets) for cell in normalized):
            return index
    return -1


def non_label_cells(row: List[str], label: str) -> List[str]:
    values: List[str] = []
    label_lower = label.lower()
    skipped = False
    for cell in row:
        normalized = normalize_inline(cell)
        if not normalized:
            continue
        if not skipped and label_lower in normalized.lower():
            skipped = True
            continue
        values.append(normalized)
    return values


def normalize_severity_text(value: str) -> str:
    cleaned = re.split(r"\s*\[", value, maxsplit=1)[0]
    return normalize_inline(cleaned)


def classify_record_status(paragraph: ParagraphBlock) -> str:
    for run in paragraph.runs:
        if not normalize_inline(run.text):
            continue
        if is_non_black_run(run):
            return "decommissioned"
    return "active"


def is_non_black_run(run: RunInfo) -> bool:
    theme_color = normalize_case(run.theme_color)
    color_value = normalize_case(run.color_value)

    if not theme_color and not color_value:
        return False
    if theme_color in BLACK_THEME_COLORS and color_value in (None, "", "auto", "000000", "000001"):
        return False
    if not theme_color and color_value in BLACK_COLOR_VALUES:
        return False
    return True


def normalize_case(value: Optional[str]) -> Optional[str]:
    if value is None:
        return None
    return value.strip().lower()


def clean_multiline_text(value: str) -> str:
    lines = [normalize_inline(line) for line in value.replace("\xa0", " ").splitlines()]
    lines = [line for line in lines if line]
    return "\n".join(lines)


def normalize_inline(value: str) -> str:
    return re.sub(r"\s+", " ", (value or "").replace("\xa0", " ")).strip()


def normalize_scalar(value: str) -> str:
    normalized = normalize_inline(value)
    if not normalized:
        return ""
    lowered = normalized.lower()
    if any(lowered.startswith(prefix) for prefix in PLACEHOLDER_PREFIXES):
        return ""
    return normalized


def normalize_block_text(value: str) -> str:
    lines = [normalize_scalar(line) for line in clean_multiline_text(value).splitlines()]
    lines = [line for line in lines if line]
    return "\n".join(lines)


def collapse_multivalue_cell(value: str) -> str:
    lines = [normalize_scalar(line) for line in clean_multiline_text(value).splitlines()]
    lines = [line for line in lines if line]
    if not lines:
        return ""
    return " | ".join(lines)


def write_csv(records: List[Dict[str, str]], output_path: Path) -> None:
    with output_path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=OUTPUT_COLUMNS)
        writer.writeheader()
        for record in records:
            writer.writerow(record)


def write_json(records: List[Dict[str, str]], output_path: Path) -> None:
    with output_path.open("w", encoding="utf-8") as handle:
        json.dump(records, handle, indent=2, ensure_ascii=False)


if __name__ == "__main__":
    try:
        main()
    except Exception as error:  # pragma: no cover
        print("Error: %s" % error, file=sys.stderr)
        sys.exit(1)
