#!/usr/bin/env python3

import argparse
import csv
import json
import os
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
W_TXBX_CONTENT = "{%s}txbxContent" % WORD_NS

SEVERITY_RE = re.compile(
    r"^\s*(?:(?P<number>\d+)\.\s*)?Severity of Breach Identified\s*:?\s*(?P<severity>.+?)\s*$",
    re.IGNORECASE,
)
SEVERITY_LABEL_ONLY_RE = re.compile(
    r"^\s*(?:(?P<number>\d+)\.\s*)?Severity of Breach Identified\s*:?\s*$",
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
    if not records:
        print(
            (
                "Warning: no breach records were found. "
                "This can happen when the document stores content inside "
                "images or unsupported Word objects."
            ),
            file=sys.stderr,
        )


def extract_breach_records(docx_path: Path) -> List[Dict[str, str]]:
    part_blocks = parse_docx_story_parts(docx_path)
    primary_blocks = part_blocks.get("word/document.xml", [])
    records = build_records_from_blocks(primary_blocks)
    if records:
        return records

    fallback_records: List[Dict[str, str]] = []
    for part_name, blocks in part_blocks.items():
        if part_name == "word/document.xml":
            continue
        fallback_records.extend(build_records_from_blocks(blocks))
    return dedupe_records(fallback_records)


def parse_docx_story_parts(docx_path: Path) -> Dict[str, List[object]]:
    with zipfile.ZipFile(str(docx_path)) as archive:
        story_parts = relevant_story_parts(archive.namelist())
        if "word/document.xml" not in story_parts:
            raise RuntimeError("word/document.xml was not found in the DOCX.")

        parsed: Dict[str, List[object]] = {}
        for part_name in story_parts:
            try:
                xml_bytes = archive.read(part_name)
            except KeyError:
                continue
            root = ET.fromstring(xml_bytes)
            body = root.find("w:body", NS)
            container = body if body is not None else root
            parsed[part_name] = collect_blocks(container)
        return parsed


def relevant_story_parts(names: List[str]) -> List[str]:
    selected: List[str] = []
    excluded_exact = {
        "word/styles.xml",
        "word/settings.xml",
        "word/fontTable.xml",
        "word/numbering.xml",
        "word/webSettings.xml",
        "word/comments.xml",
        "word/commentsExtended.xml",
        "word/commentsIds.xml",
        "word/theme/theme1.xml",
    }
    excluded_prefixes = (
        "word/_rels/",
        "word/theme/",
    )

    for name in sorted(names):
        if not name.startswith("word/") or not name.endswith(".xml"):
            continue
        if name in excluded_exact:
            continue
        if any(name.startswith(prefix) for prefix in excluded_prefixes):
            continue
        basename = os.path.basename(name)
        if basename in {"app.xml", "core.xml"}:
            continue
        selected.append(name)
    return selected


def collect_blocks(container: ET.Element) -> List[object]:
    blocks: List[object] = []
    for child in list(container):
        if child.tag == W_TBL:
            blocks.extend(collect_blocks_from_table(child))
            continue

        if child.tag == W_P:
            paragraph = parse_paragraph_block(child)
            blocks.append(paragraph)
            blocks.extend(extract_textbox_blocks(child))
            continue

        blocks.extend(collect_blocks(child))
    return blocks


def collect_blocks_from_table(table_element: ET.Element) -> List[object]:
    blocks: List[object] = []
    for row_element in table_element.findall("./w:tr", NS):
        for cell_element in row_element.findall("./w:tc", NS):
            for child in list(cell_element):
                if child.tag == W_P:
                    paragraph = parse_paragraph_block(child)
                    blocks.append(paragraph)
                    blocks.extend(extract_textbox_blocks(child))
                    continue
                if child.tag == W_TBL:
                    blocks.extend(collect_blocks_from_table(child))
                    continue
                blocks.extend(collect_blocks(child))

    blocks.append(parse_table_block(table_element))
    return blocks


def extract_textbox_blocks(paragraph_element: ET.Element) -> List[object]:
    blocks: List[object] = []
    for textbox_content in paragraph_element.iter(W_TXBX_CONTENT):
        blocks.extend(collect_blocks(textbox_content))
    return blocks


def parse_paragraph_block(paragraph_element: ET.Element) -> ParagraphBlock:
    runs: List[RunInfo] = []
    parts: List[str] = []
    for run_element in iter_paragraph_runs(paragraph_element):
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


def iter_paragraph_runs(paragraph_element: ET.Element) -> Iterable[ET.Element]:
    for child in list(paragraph_element):
        if child.tag == W_R:
            yield child
            continue
        if child.tag in (W_P, W_TBL):
            continue
        for run_element in iter_paragraph_runs(child):
            yield run_element


def parse_table_block(table_element: ET.Element) -> TableBlock:
    rows: List[List[str]] = []
    grid_width = detect_table_grid_width(table_element)
    for row_element in table_element.findall("./w:tr", NS):
        row = parse_table_row(row_element, grid_width)
        if any(normalize_inline(cell) for cell in row):
            rows.append(row)
    return TableBlock(rows=rows)


def detect_table_grid_width(table_element: ET.Element) -> int:
    grid_columns = table_element.findall("./w:tblGrid/w:gridCol", NS)
    if grid_columns:
        return len(grid_columns)

    estimated_width = 0
    for row_element in table_element.findall("./w:tr", NS):
        estimated_width = max(estimated_width, estimate_row_grid_width(row_element))
    return estimated_width


def estimate_row_grid_width(row_element: ET.Element) -> int:
    width = extract_row_grid_offset(row_element, "gridBefore")
    for cell_element in row_element.findall("./w:tc", NS):
        width += extract_grid_span(cell_element)
    width += extract_row_grid_offset(row_element, "gridAfter")
    return width


def parse_table_row(row_element: ET.Element, grid_width: int) -> List[str]:
    row: List[str] = [""] * extract_row_grid_offset(row_element, "gridBefore")

    for cell_element in row_element.findall("./w:tc", NS):
        cell_text = extract_cell_text(cell_element)
        span = extract_grid_span(cell_element)
        row.append(cell_text)
        if span > 1:
            row.extend([""] * (span - 1))

    row.extend([""] * extract_row_grid_offset(row_element, "gridAfter"))
    if grid_width and len(row) < grid_width:
        row.extend([""] * (grid_width - len(row)))
    return row


def extract_row_grid_offset(row_element: ET.Element, tag_name: str) -> int:
    offset_element = row_element.find("./w:trPr/w:%s" % tag_name, NS)
    if offset_element is None:
        return 0
    return parse_int_value(offset_element.get(W_VAL), default=0)


def extract_grid_span(cell_element: ET.Element) -> int:
    span_element = cell_element.find("./w:tcPr/w:gridSpan", NS)
    if span_element is None:
        return 1
    return max(1, parse_int_value(span_element.get(W_VAL), default=1))


def extract_cell_text(cell_element: ET.Element) -> str:
    paragraphs = []
    for paragraph_element in cell_element.findall(".//w:p", NS):
        paragraph = parse_paragraph_block(paragraph_element)
        if paragraph.text:
            paragraphs.append(paragraph.text)
    return "\n".join(paragraphs).strip()


def parse_int_value(value: Optional[str], default: int) -> int:
    if value is None:
        return default
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


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
    chunk_sequence = 0
    pending_severity: Optional[Dict[str, str]] = None

    def make_chunk_number(explicit_number: Optional[str]) -> str:
        nonlocal chunk_sequence
        if explicit_number:
            try:
                chunk_sequence = max(chunk_sequence, int(explicit_number))
            except ValueError:
                pass
            return explicit_number

        chunk_sequence += 1
        return str(chunk_sequence)

    def start_chunk(
        explicit_number: Optional[str],
        severity_text: str,
        record_status: str,
        table_rows: Optional[List[List[str]]] = None,
    ) -> None:
        nonlocal current_chunk
        if current_chunk is not None:
            records.append(finalize_chunk(current_chunk))
        current_chunk = {
            "group_name": current_group,
            "chunk_number": make_chunk_number(explicit_number),
            "severity_raw": normalize_severity_text(severity_text),
            "record_status": record_status,
            "table": table_rows,
        }

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

            label_only_match = SEVERITY_LABEL_ONLY_RE.match(text)
            if label_only_match:
                pending_severity = {
                    "number": label_only_match.group("number") or "",
                    "record_status": classify_record_status(block),
                }
                continue

            if pending_severity is not None and looks_like_severity_value(text):
                start_chunk(
                    pending_severity.get("number") or None,
                    text,
                    pending_severity.get("record_status") or "active",
                    None,
                )
                pending_severity = None
                continue

            severity_match = SEVERITY_RE.match(text)
            if severity_match:
                start_chunk(
                    severity_match.group("number"),
                    severity_match.group("severity"),
                    classify_record_status(block),
                    None,
                )
                pending_severity = None
                continue

            if GROUP_RE.match(text):
                current_group = text
                pending_severity = None
                continue

        elif isinstance(block, TableBlock):
            row_texts = [normalize_inline(" ".join(row)) for row in block.rows]

            for row_text in row_texts:
                if not row_text:
                    continue

                if not in_breach_section and "performance breaches:" in row_text.lower():
                    in_breach_section = True

                if not in_breach_section:
                    continue

                if GROUP_RE.match(row_text):
                    current_group = row_text
                    pending_severity = None
                    continue

                severity_match = SEVERITY_RE.match(row_text)
                if severity_match:
                    severity_text = normalize_severity_text(severity_match.group("severity"))
                    severity_number = severity_match.group("number")
                    if (
                        current_chunk is not None
                        and current_chunk.get("table") is None
                        and current_chunk.get("severity_raw") == severity_text
                        and (
                            not severity_number
                            or str(current_chunk.get("chunk_number", "")) == severity_number
                        )
                    ):
                        current_chunk["table"] = block.rows
                        pending_severity = None
                        continue

                    start_chunk(
                        severity_number,
                        severity_text,
                        (pending_severity or {}).get("record_status", "active"),
                        block.rows,
                    )
                    pending_severity = None

            if in_breach_section and current_chunk is not None and current_chunk.get("table") is None:
                current_chunk["table"] = block.rows

    if current_chunk is not None:
        records.append(finalize_chunk(current_chunk))

    return records


def dedupe_records(records: List[Dict[str, str]]) -> List[Dict[str, str]]:
    unique: List[Dict[str, str]] = []
    seen = set()
    for record in records:
        key = (
            record.get("group_name", ""),
            record.get("chunk_number", ""),
            record.get("severity_raw", ""),
            record.get("model_ids", ""),
            record.get("model_names", ""),
        )
        if key in seen:
            continue
        seen.add(key)
        unique.append(record)
    return unique


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

    model_header_index = find_model_header_row_index(rows)
    model_column_map = {}
    if model_header_index != -1:
        model_column_map = build_model_column_map(rows[model_header_index])

    for row_index, row in enumerate(rows):
        row_lower = [normalize_inline(cell).lower() for cell in row]

        if any("model(s) affected" in cell for cell in row_lower):
            extracted.update(
                extract_model_affected_values(
                    row=row,
                    row_index=row_index,
                    model_column_map=model_column_map,
                )
            )

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


def find_model_header_row_index(rows: List[List[str]]) -> int:
    required = ("model risk rating", "model id", "model name")
    for index, row in enumerate(rows):
        normalized = [normalize_inline(cell).lower() for cell in row]
        if all(any(required_text in cell for cell in normalized) for required_text in required):
            return index
    return -1


def build_model_column_map(header_row: List[str]) -> Dict[str, int]:
    column_map: Dict[str, int] = {}
    for index, cell in enumerate(header_row):
        label = normalize_inline(cell).lower()
        if "model risk rating" in label:
            column_map["model_risk_rating"] = index
        elif "model id" in label:
            column_map["model_ids"] = index
        elif "model name" in label:
            column_map["model_names"] = index
    return column_map


def extract_model_affected_values(
    row: List[str], row_index: int, model_column_map: Dict[str, int]
) -> Dict[str, str]:
    extracted = {
        "model_risk_rating": "",
        "model_ids": "",
        "model_names": "",
    }
    embedded_risk_rating = extract_embedded_label_value(row, "model(s) affected")

    # Preferred path: use the actual header columns from the table.
    if model_column_map:
        for field_name, column_index in model_column_map.items():
            if column_index < len(row):
                value = row[column_index]
                if field_name == "model_risk_rating":
                    extracted[field_name] = normalize_scalar(value)
                else:
                    extracted[field_name] = collapse_multivalue_cell(value)

        if not extracted["model_risk_rating"] and embedded_risk_rating:
            extracted["model_risk_rating"] = normalize_scalar(embedded_risk_rating)

        if any(extracted.values()):
            return extracted

    # Conservative fallback: only populate when the values clearly look like
    # the A-layout columns. For B-layout legacy tables we leave these blank.
    values = []
    if embedded_risk_rating:
        values.append(embedded_risk_rating)
    values.extend(non_label_cells(row, "model(s) affected"))

    if len(values) >= 3 and looks_like_risk_rating(values[0]) and looks_like_model_id(values[1]):
        extracted["model_risk_rating"] = normalize_scalar(values[0])
        extracted["model_ids"] = collapse_multivalue_cell(values[1])
        extracted["model_names"] = collapse_multivalue_cell(values[2])
    elif len(values) >= 2 and looks_like_model_id(values[0]):
        extracted["model_ids"] = collapse_multivalue_cell(values[0])
        extracted["model_names"] = collapse_multivalue_cell(values[1])

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


def extract_embedded_label_value(row: List[str], label: str) -> str:
    label_lower = label.lower()
    for cell in row:
        cleaned = clean_multiline_text(cell)
        if not cleaned:
            continue
        lines = cleaned.splitlines()
        first_line = normalize_inline(lines[0])
        if label_lower not in first_line.lower():
            continue

        residual_lines: List[str] = []
        first_line_without_label = re.sub(re.escape(label), "", first_line, flags=re.IGNORECASE)
        first_line_without_label = first_line_without_label.strip(" :\t")
        if first_line_without_label:
            residual_lines.append(first_line_without_label)
        residual_lines.extend(lines[1:])
        residual_lines = [normalize_scalar(line) for line in residual_lines]
        residual_lines = [line for line in residual_lines if line]
        if residual_lines:
            return " | ".join(residual_lines)
    return ""


def normalize_severity_text(value: str) -> str:
    cleaned = re.split(r"\s*\[", value, maxsplit=1)[0]
    return normalize_inline(cleaned)


def looks_like_severity_value(value: str) -> bool:
    normalized = normalize_inline(value)
    lowered = normalized.lower()
    severity_keywords = (
        "persistent breach",
        "exception required",
        "minor",
        "temporary condition",
    )
    return len(normalized) <= 120 and any(keyword in lowered for keyword in severity_keywords)


def looks_like_risk_rating(value: str) -> bool:
    lowered = normalize_inline(value).lower()
    return lowered in {"low", "medium", "high"}


def looks_like_model_id(value: str) -> bool:
    normalized = normalize_inline(value).upper()
    return normalized.startswith("MOD_")


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
    with output_path.open("w", encoding="utf-8-sig", newline="") as handle:
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
