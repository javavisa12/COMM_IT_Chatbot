"""Read uploaded files, extract text, and send valid content to indexing."""

import csv
import io
import os
import posixpath
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter

from fastapi import UploadFile
from pypdf import PdfReader
import pymupdf
import pandas as pd
import xlrd

from backend import core
from backend.services.rag_service import (
    index_file,
    restore_persisted_session,
    search_context
)


MAXIMO_HEADERLESS_CSV_LABELS = [
    "Priority",
    "Description",
    "Location",
    "Is dynamic",
    "Work Order",
    "Site",
    "Organization",
    "Column 8",
    "Scheduled Start",
    "Asset",
    "Status"
]

STATUS_ALIASES = {
    "approved": {"approved", "appr"},
    "not approved": {"not approved", "nappr"},
    "revised": {"revised", "revisd"},
    "waiting approval": {"waiting approval", "waiting for approval", "wappr"},
    "waiting materials": {"waiting materials", "wmatl"},
    "waiting to be scheduled": {"waiting to be scheduled", "wsch"},
    "in progress": {"in progress", "inprg"}
}

STATUS_DESCRIPTIONS = {
    "approved": "approved",
    "not approved": "not approved",
    "waiting approval": "waiting for approval",
    "waiting materials": "waiting for materials",
    "revised": "revised",
    "waiting to be scheduled": "waiting to be scheduled",
    "in progress": "in progress"
}


def _status_matches_question(status, question):

    status_lower = status.strip().lower()
    question_lower = (question or "").lower()
    aliases = {status_lower}

    for label, values in STATUS_ALIASES.items():
        if status_lower in values:
            aliases.update(values)
            aliases.add(label)

    return any(
        re.search(rf"\b{re.escape(alias)}\b", question_lower)
        for alias in aliases
    )


def _status_display_name(status, question):

    status_lower = status.strip().lower()
    for label, values in STATUS_ALIASES.items():
        if status_lower in values:
            return label

    return status


def _xlsx_shared_strings(zip_file):

    if "xl/sharedStrings.xml" not in zip_file.namelist():
        return []

    try:
        root = ET.fromstring(zip_file.read("xl/sharedStrings.xml"))
        ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        values = []

        for item in root.findall("main:si", ns):
            text = "".join(
                node.text or ""
                for node in item.findall(".//main:t", ns)
            )
            values.append(text)

        return values

    except Exception as error:
        print(f"SHARED STRINGS PARSE FAILED: {error}")
        return []


def _xlsx_cell_value(cell, shared_strings):

    cell_type = cell.attrib.get("t")
    ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

    if cell_type == "inlineStr":
        return "".join(
            node.text or ""
            for node in cell.findall(".//main:t", ns)
        )

    if cell_type == "s":
        value_node = cell.find("main:v", ns)
        if value_node is None or value_node.text is None:
            return ""
        try:
            index = int(value_node.text)
            if 0 <= index < len(shared_strings):
                return shared_strings[index]
        except ValueError:
            pass
        return ""

    value_node = cell.find("main:v", ns)
    if value_node is None or value_node.text is None:
        return ""
    return value_node.text


def _xlsx_column_index(cell_reference):

    column_letters = "".join(
        character
        for character in cell_reference
        if character.isalpha()
    )
    column_index = 0

    for character in column_letters:
        column_index = column_index * 26 + ord(character.upper()) - ord("A") + 1

    return column_index - 1


def _extract_excel_text(filepath):

    try:
        if filepath.lower().endswith((".xlsx", ".xlsm")):
            with zipfile.ZipFile(filepath) as archive:
                shared_strings = _xlsx_shared_strings(archive)

                workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
                rels_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
                workbook_ns = {
                    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                    "pkg": "http://schemas.openxmlformats.org/package/2006/relationships"
                }

                rel_map = {
                    relation.attrib.get("Id"): relation.attrib.get("Target", "")
                    for relation in rels_root.findall("pkg:Relationship", workbook_ns)
                }

                text_parts = []
                for sheet in workbook_root.findall("main:sheets/main:sheet", workbook_ns):
                    rel_id = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                    target = rel_map.get(rel_id, "")
                    if not target:
                        continue

                    if target.startswith("/"):
                        target = target[1:]
                    if not target.startswith("xl/"):
                        target = posixpath.join("xl", target)
                    target = posixpath.normpath(target)

                    if target not in archive.namelist():
                        continue

                    sheet_root = ET.fromstring(archive.read(target))
                    rows = []
                    for row in sheet_root.findall(".//main:sheetData/main:row", workbook_ns):
                        values = []
                        for cell in row.findall("main:c", workbook_ns):
                            value = _xlsx_cell_value(cell, shared_strings)
                            cell_index = _xlsx_column_index(cell.attrib.get("r", ""))
                            while len(values) <= cell_index:
                                values.append("")
                            values[cell_index] = value.strip()

                        while values and not values[-1]:
                            values.pop()
                        if any(values):
                            rows.append(" | ".join(values))

                    if rows:
                        text_parts.append("\n".join(rows))

                return "\n\n".join(text_parts)

        if filepath.lower().endswith(".xls"):
            workbook = xlrd.open_workbook(filepath)
            text_parts = []
            for sheet_name in workbook.sheet_names():
                sheet = workbook.sheet_by_name(sheet_name)
                rows = []
                for row_index in range(sheet.nrows):
                    values = [str(value).strip() for value in sheet.row_values(row_index)]
                    cleaned = [value for value in values if value]
                    if cleaned:
                        rows.append(" | ".join(cleaned))
                if rows:
                    text_parts.append("\n".join(rows))
            return "\n\n".join(text_parts)

    except Exception as error:
        print(f"EXCEL EXTRACTION FAILED: {error}")

    return ""


def load_file(filepath):
    """Extract readable text from PDF, spreadsheet, CSV, or text files."""

    try:
        lower_filepath = filepath.lower()

        if lower_filepath.endswith(".pdf"):
            try:
                reader = PdfReader(filepath)
                text = "\n\n".join(
                    page.extract_text() or ""
                    for page in reader.pages
                )
                if text.strip():
                    return text
            except Exception as error:
                print(f"PYPDF EXTRACTION FAILED: {error}")

            document = pymupdf.open(filepath)
            text = "\n\n".join(page.get_text() for page in document)
            if text.strip():
                print("PDF EXTRACTED WITH PYMUPDF:", filepath)
                return text

            print("PDF CONTAINS NO EXTRACTABLE TEXT:", filepath)
            return ""

        if lower_filepath.endswith((".xlsx", ".xlsm", ".xls")):
            text = _extract_excel_text(filepath)
            if text and text.strip():
                return text
            print("EXCEL CONTAINS NO EXTRACTABLE TEXT:", filepath)
            return ""

        if lower_filepath.endswith(".csv"):
            with open(
                filepath,
                "r",
                encoding="utf-8",
                errors="ignore"
            ) as file:
                return file.read()

        with open(
            filepath,
            "r",
            encoding="utf-8",
            errors="ignore"
        ) as file:
            return file.read()

    except Exception as error:
        print(f"ERROR READING FILE: {error}")
        return ""

def _spreadsheet_rows(filepath):
    """Return spreadsheet rows, supporting both XLSX-style and CSV data."""

    text = load_file(filepath)
    if filepath.lower().endswith(".csv"):
        return [
            [value.strip() for value in row]
            for row in csv.reader(io.StringIO(text))
            if any(value.strip() for value in row)
        ]

    return [
        line.split(" | ")
        for line in text.splitlines()
        if line.strip()
    ]


def _normalized_tabular_records(filepath):
    """Return labeled tabular row documents and normalized column metadata."""

    if not filepath.lower().endswith((".csv", ".xlsx", ".xlsm", ".xls")):
        return None

    try:
        headerless_csv = False
        if filepath.lower().endswith(".csv"):
            with open(
                filepath,
                "r",
                encoding="utf-8",
                errors="ignore"
            ) as file:
                sample = file.readline().lower()
            header_keywords = {
                "asset", "description", "location", "work order", "status"
            }
            has_header = any(keyword in sample for keyword in header_keywords)
            headerless_csv = not has_header
            frame = pd.read_csv(
                filepath,
                header=0 if has_header else None,
                dtype=str,
                keep_default_na=False
            )
        else:
            frame = pd.read_excel(
                filepath,
                dtype=str,
                keep_default_na=False
            )

        if frame.empty:
            return []

        if headerless_csv:
            frame.columns = [
                MAXIMO_HEADERLESS_CSV_LABELS[index]
                if index < len(MAXIMO_HEADERLESS_CSV_LABELS)
                else f"Column {index + 1}"
                for index in range(len(frame.columns))
            ]

        frame.columns = [
            str(column).strip() or f"Column {index + 1}"
            for index, column in enumerate(frame.columns)
        ]
        boolean_value_map = {
            "y": "Yes",
            "yes": "Yes",
            "true": "Yes",
            "n": "No",
            "no": "No",
            "false": "No"
        }
        for column in frame.columns:
            values = {
                str(value).strip().casefold()
                for value in frame[column]
                if str(value).strip()
            }
            if values and values.issubset(boolean_value_map):
                frame[column] = frame[column].map(
                    lambda value: boolean_value_map.get(
                        str(value).strip().casefold(), ""
                    )
                )

        column_keys = {
            re.sub(r"[^a-z0-9]+", "_", str(column).lower()).strip("_")
            for column in frame.columns
        }
        if "work_order" in column_keys:
            dataset_type = "work_order"
        elif "item" in column_keys or "item_number" in column_keys:
            dataset_type = "inventory"
        elif "asset" in column_keys:
            dataset_type = "asset"
        else:
            dataset_type = "tabular"

        column_names = "|".join(str(column) for column in frame.columns)
        documents = []
        row_metadata = []
        for row_index, (_, row) in enumerate(frame.iterrows()):
            attributes = [
                f"{column}={str(value).strip()}"
                for column, value in row.items()
                if str(value).strip()
            ]
            if attributes:
                documents.append(
                    f"Row {row_index + 2}: " + " | ".join(attributes)
                )
                metadata = {
                    "row_index": row_index,
                    "dataset_type": dataset_type,
                    "column_names": column_names
                }
                for column, value in row.items():
                    value = str(value).strip()
                    column_key = re.sub(
                        r"[^a-z0-9]+",
                        "_",
                        str(column).strip().lower()
                    ).strip("_")
                    if value and column_key:
                        metadata[f"attr_{column_key}"] = value
                row_metadata.append(metadata)
        return documents, row_metadata

    except Exception as error:
        print(f"PANDAS NORMALIZATION FAILED: {error}")
        return None


def question_references_tabular_column(question):
    """Detect whether a question names a column in an active spreadsheet."""

    normalized_question = re.sub(r"[^a-z0-9]+", " ", (question or "").lower())
    restore_persisted_session()

    for filepath in core.SESSION_FILES:
        normalized_records = _normalized_tabular_records(filepath)
        if not normalized_records or not normalized_records[1]:
            continue

        for column in normalized_records[1][0].get("column_names", "").split("|"):
            normalized_column = re.sub(r"[^a-z0-9]+", " ", column.lower()).strip()
            if not normalized_column:
                continue
            if re.search(rf"\b{re.escape(normalized_column)}s?\b", normalized_question):
                return True

    return False


def _spreadsheet_context(filepath):
    """Format every spreadsheet row with column labels for the LLM prompt."""

    rows = _spreadsheet_rows(filepath)
    if not rows:
        return ""

    header = rows[0]
    normalized_header = [column.strip().lower() for column in header]
    has_header = any(
        column in {
            "work order", "description", "location", "asset", "status",
            "scheduled start", "priority", "site"
        }
        for column in normalized_header
    )
    if has_header:
        labels = [
            column or f"Column {index + 1}"
            for index, column in enumerate(header)
        ]
        data_rows = rows[1:]
    else:
        labels = [
            MAXIMO_HEADERLESS_CSV_LABELS[index]
            if index < len(MAXIMO_HEADERLESS_CSV_LABELS)
            else f"Column {index + 1}"
            for index in range(max(map(len, rows)))
        ]
        data_rows = rows

    formatted_rows = []
    for row_number, row in enumerate(data_rows, start=2 if has_header else 1):
        attributes = [
            f"{labels[index]}={value}"
            for index, value in enumerate(row)
            if value.strip()
        ]
        if attributes:
            formatted_rows.append(
                f"Row {row_number}: " + " | ".join(attributes)
            )

    return "\n".join(formatted_rows)


def _is_work_order_count_question(question):

    text = (question or "").lower()
    if not re.search(r"\b(how many|how much|count|total|number of)\b", text):
        return False

    work_order_keywords = [
        "work order",
        "work orders",
        "status",
        "open",
        "closed",
        "site",
        "contract",
        "annual leave",
        "maximo",
        "request",
        "job"
    ]

    return any(keyword in text for keyword in work_order_keywords)


def build_file_context(question):
    """Use exact column filters first, then semantic search for file context."""
    return search_context(question)


def build_spreadsheet_count_answer(question):
    """Answer supported spreadsheet count questions directly from all rows."""

    if not re.search(r"\b(how many|how much|count|total|number of)\b", (question or "").lower()):
        return ""

    restore_persisted_session()
    answers = []
    is_work_order_question = _is_work_order_count_question(question)
    aggregated_status_counts = Counter()

    for filepath in core.SESSION_FILES:
        if not filepath.lower().endswith((".xlsx", ".xlsm", ".xls", ".csv")):
            continue

        rows = _spreadsheet_rows(filepath)
        if not rows:
            continue

        header = rows[0]
        normalized_header = [column.strip().lower() for column in header]
        status_index = next(
            (
                index
                for index, column in enumerate(normalized_header)
                if column in {"status", "approval status", "approval"}
            ),
            None
        )
        work_order_index = next(
            (
                index
                for index, column in enumerate(normalized_header)
                if column == "work order"
            ),
            None
        )
        has_header = status_index is not None or work_order_index is not None
        data_rows = rows[1:] if has_header else rows

        if status_index is None:
            if not filepath.lower().endswith(".csv") or not header:
                continue
            status_index = len(header) - 1

        if work_order_index is None:
            status_counts = Counter(
                row[status_index].strip()
                for row in data_rows
                if len(row) > status_index and row[status_index].strip()
            )
            question_lower = question.lower()
            requested_statuses = [
                status
                for status in status_counts
                if _status_matches_question(status, question)
            ]
            if not requested_statuses:
                continue

            filename = os.path.basename(filepath)
            first_status = requested_statuses[0]
            first_status_label = _status_display_name(first_status, question)
            status_text = ", ".join(
                f"Status {_status_display_name(status, question)}: {status_counts[status]} records"
                for status in requested_statuses
            )
            if len(core.SESSION_FILES) > 1:
                for status in requested_statuses:
                    aggregated_status_counts[status] += status_counts[status]
                continue

            answers.append(
                f"There are {status_counts[first_status]} {first_status_label.lower()} records in {filename}. "
                f"{status_text}. [Source: {filename}, spreadsheet summary]"
            )
            continue

        if not is_work_order_question:
            status_index = rows[0].index("Status")
            status_values = {
                row[status_index].strip()
                for row in rows[1:]
                if len(row) > status_index and row[status_index].strip()
            }
            is_status_question = any(
                _status_matches_question(status, question)
                for status in status_values
            )
            if not is_status_question:
                continue

        header = rows[0]
        work_order_index = header.index("Work Order")
        status_index = header.index("Status")
        work_orders = [
            row[work_order_index]
            for row in rows[1:]
            if len(row) > work_order_index and row[work_order_index].strip()
        ]
        statuses = [
            row[status_index]
            for row in rows[1:]
            if len(row) > status_index and row[status_index].strip()
        ]
        status_counts = Counter(statuses)
        filename = os.path.basename(filepath)
        question_lower = question.lower()
        requested_statuses = [
            status
            for status in status_counts
            if _status_matches_question(status, question)
        ]

        if len(core.SESSION_FILES) > 1 and requested_statuses:
            for status in requested_statuses:
                aggregated_status_counts[status] += status_counts[status]
            continue

        answer = (
            f"There are {len(work_orders)} work orders in {filename}."
        )
        for status in requested_statuses:
            status_label = _status_display_name(status, question)
            answer += (
                f" Status {status_label}: {status_counts[status]} work orders"
                f" ({status_counts[status]} {status_label.lower()})."
            )
        answer += f" [Source: {filename}, spreadsheet summary]"
        answers.append(answer)

    if len(core.SESSION_FILES) > 1 and aggregated_status_counts:
        total_count = sum(aggregated_status_counts.values())
        labels = []
        for status in sorted(aggregated_status_counts):
            status_label = _status_display_name(status, question)
            labels.append(
                f"Status {status_label}: {aggregated_status_counts[status]} work orders"
                f" ({aggregated_status_counts[status]} {status_label.lower()})."
            )
        return (
            f"There are {total_count} matching work orders across the active files. "
            + " ".join(labels)
            + " [Source: active files, spreadsheet summary]"
        )

    return "\n\n".join(answers)


def build_spreadsheet_location_answer(question):
    """Return the most frequently used location in uploaded spreadsheets."""

    question_lower = (question or "").lower()
    if "location" not in question_lower:
        return ""
    if not re.search(r"\b(most|highest|common|frequent|used)\b", question_lower):
        return ""

    restore_persisted_session()
    answers = []

    for filepath in core.SESSION_FILES:
        if not filepath.lower().endswith((".xlsx", ".xlsm", ".xls", ".csv")):
            continue

        rows = _spreadsheet_rows(filepath)
        if not rows:
            continue

        header = rows[0]
        normalized_header = [column.strip().lower() for column in header]
        location_index = next(
            (
                index
                for index, column in enumerate(normalized_header)
                if column in {"location", "locations"}
            ),
            None
        )
        has_header = location_index is not None

        if location_index is None:
            if not filepath.lower().endswith(".csv") or len(header) <= 2:
                continue
            location_index = 2

        locations = [
            row[location_index].strip()
            for row in (rows[1:] if has_header else rows)
            if len(row) > location_index and row[location_index].strip()
        ]
        if not locations:
            continue

        location, count = Counter(locations).most_common(1)[0]
        filename = os.path.basename(filepath)
        answers.append(
            f"The most used location is {location}, with {count} work orders "
            f"in {filename}. [Source: {filename}, spreadsheet summary]"
        )

    return "\n\n".join(answers)


def build_spreadsheet_location_count_answer(question):
    """Count rows whose spreadsheet location matches the requested value."""

    question_lower = (question or "").lower()
    if "location" not in question_lower:
        return ""
    if not re.search(r"\b(how many|count|number of|total)\b", question_lower):
        return ""

    location_match = re.search(
        r"\blocation\s+(?:is|=)?\s*['\"]?([\w-]+)",
        question,
        re.IGNORECASE
    )
    if not location_match:
        return ""

    requested_location = location_match.group(1).strip()
    restore_persisted_session()
    answers = []

    for filepath in core.SESSION_FILES:
        if not filepath.lower().endswith((".xlsx", ".xlsm", ".xls", ".csv")):
            continue

        rows = _spreadsheet_rows(filepath)
        if not rows:
            continue

        header = rows[0]
        normalized_header = [column.strip().lower() for column in header]
        location_index = next(
            (
                index
                for index, column in enumerate(normalized_header)
                if column in {"location", "locations"}
            ),
            None
        )
        has_header = location_index is not None

        if location_index is None:
            if not filepath.lower().endswith(".csv") or len(header) <= 2:
                continue
            location_index = 2

        matching_rows = sum(
            1
            for row in (rows[1:] if has_header else rows)
            if len(row) > location_index
            and row[location_index].strip().lower() == requested_location.lower()
        )
        if matching_rows == 0:
            continue

        filename = os.path.basename(filepath)
        answers.append(
            f"The location {requested_location.upper()} appears {matching_rows} "
            f"times in {filename}. [Source: {filename}, spreadsheet summary]"
        )

    return "\n\n".join(answers)


def build_spreadsheet_distinct_location_count_answer(question):
    """Count unique non-empty values in the spreadsheet Location column."""

    question_lower = (question or "").lower()
    if "location" not in question_lower:
        return ""
    if not re.search(r"\b(different|distinct|unique)\b", question_lower):
        return ""
    if not re.search(r"\b(how many|count|number of|total)\b", question_lower):
        return ""

    restore_persisted_session()
    answers = []

    for filepath in core.SESSION_FILES:
        if not filepath.lower().endswith((".xlsx", ".xlsm", ".xls", ".csv")):
            continue

        rows = _spreadsheet_rows(filepath)
        if not rows:
            continue

        header = rows[0]
        normalized_header = [column.strip().lower() for column in header]
        location_index = next(
            (
                index
                for index, column in enumerate(normalized_header)
                if column in {"location", "locations"}
            ),
            None
        )
        has_header = location_index is not None

        if location_index is None:
            if not filepath.lower().endswith(".csv") or len(header) <= 2:
                continue
            location_index = 2

        locations = {
            row[location_index].strip()
            for row in (rows[1:] if has_header else rows)
            if len(row) > location_index and row[location_index].strip()
        }
        if not locations:
            continue

        filename = os.path.basename(filepath)
        answers.append(
            f"There are {len(locations)} different locations in {filename}. "
            f"[Source: {filename}, spreadsheet summary]"
        )

    return "\n\n".join(answers)


def build_spreadsheet_distinct_column_count_answer(question):
    """Count unique non-empty values in a requested spreadsheet column."""

    question_lower = re.sub(r"[^a-z0-9]+", " ", (question or "").lower())
    if not re.search(r"\b(different|distinct|unique)\b", question_lower):
        return ""
    if not re.search(r"\b(how many|count|number of|total)\b", question_lower):
        return ""

    restore_persisted_session()
    answers = []

    for filepath in core.SESSION_FILES:
        if not filepath.lower().endswith((".xlsx", ".xlsm", ".xls", ".csv")):
            continue

        normalized_records = _normalized_tabular_records(filepath)
        if not normalized_records:
            continue

        row_metadata = normalized_records[1]
        if not row_metadata:
            continue

        column_names = row_metadata[0].get("column_names", "").split("|")
        requested_column = next(
            (
                column
                for column in column_names
                if re.sub(r"[^a-z0-9]+", " ", column.lower()).strip()
                in question_lower
            ),
            None
        )
        if not requested_column:
            continue

        column_key = re.sub(
            r"[^a-z0-9]+",
            "_",
            requested_column.lower()
        ).strip("_")
        values = {
            metadata[f"attr_{column_key}"].strip().casefold()
            for metadata in row_metadata
            if metadata.get(f"attr_{column_key}", "").strip()
        }
        if not values:
            continue

        filename = os.path.basename(filepath)
        answers.append(
            f"There are {len(values)} different {requested_column} values "
            f"in {filename}. [Source: {filename}, spreadsheet summary]"
        )

    return "\n\n".join(answers)


def build_spreadsheet_column_count_answer(question):
    """Count distinct column values or rows matching a requested column value."""

    question_lower = re.sub(r"[^a-z0-9]+", " ", (question or "").lower())
    if not re.search(r"\b(how many|count|number of|total)\b", question_lower):
        return ""

    restore_persisted_session()
    answers = []

    for filepath in core.SESSION_FILES:
        normalized_records = _normalized_tabular_records(filepath)
        if not normalized_records or not normalized_records[1]:
            continue

        row_metadata = normalized_records[1]
        column_names = row_metadata[0].get("column_names", "").split("|")
        requested_column = next(
            (
                column
                for column in column_names
                if re.search(
                    rf"\b{re.escape(re.sub(r'[^a-z0-9]+', ' ', column.lower()).strip())}s?\b",
                    question_lower
                )
            ),
            None
        )
        if not requested_column:
            continue

        column_key = re.sub(
            r"[^a-z0-9]+",
            "_",
            requested_column.lower()
        ).strip("_")
        column_values = [
            metadata.get(f"attr_{column_key}", "").strip()
            for metadata in row_metadata
            if metadata.get(f"attr_{column_key}", "").strip()
        ]
        if not column_values:
            continue

        matching_values = {
            value.casefold()
            for value in column_values
            if (
                _status_matches_question(value, question)
                if column_key == "status"
                else re.sub(r"[^a-z0-9]+", " ", value.lower()).strip()
                in question_lower
            )
        }
        if matching_values:
            count = sum(
                value.casefold() in matching_values
                for value in column_values
            )
            display_values = {
                _status_display_name(value, question)
                if column_key == "status"
                else value
                for value in column_values
                if value.casefold() in matching_values
            }
            answer = (
                f"There are {count} records with {requested_column} "
                f"{', '.join(sorted(display_values, key=str.casefold))}"
            )
        else:
            distinct_values = {value.casefold() for value in column_values}
            count = len(distinct_values)
            answer = (
                f"There are {count} different {requested_column} values"
            )

        filename = os.path.basename(filepath)
        answers.append(
            f"{answer} in {filename}. "
            f"[Source: {filename}, spreadsheet summary]"
        )

    return "\n\n".join(answers)


def build_spreadsheet_asset_count_answer(question):
    """Count distinct non-empty asset values in uploaded spreadsheets."""

    question_lower = (question or "").lower()
    if "asset" not in question_lower:
        return ""
    if not re.search(r"\b(how many|count|number of|total)\b", question_lower):
        return ""

    restore_persisted_session()
    answers = []

    for filepath in core.SESSION_FILES:
        if not filepath.lower().endswith((".xlsx", ".xlsm", ".xls", ".csv")):
            continue

        rows = _spreadsheet_rows(filepath)
        if not rows:
            continue

        header = rows[0]
        normalized_header = [column.strip().lower() for column in header]
        asset_index = next(
            (
                index
                for index, column in enumerate(normalized_header)
                if column == "asset"
            ),
            None
        )
        has_header = asset_index is not None

        if asset_index is None:
            if not filepath.lower().endswith(".csv") or len(header) <= 9:
                continue
            asset_index = 9

        assets = {
            row[asset_index].strip()
            for row in (rows[1:] if has_header else rows)
            if len(row) > asset_index and row[asset_index].strip()
        }
        if not assets:
            continue

        filename = os.path.basename(filepath)
        answers.append(
            f"There are {len(assets)} distinct assets in {filename}. "
            f"[Source: {filename}, spreadsheet summary]"
        )

    return "\n\n".join(answers)


def build_spreadsheet_item_count_answer(question):
    """Count distinct inventory items in spreadsheets with an Item column."""

    question_lower = (question or "").lower()
    if not re.search(r"\b(item|items|inventory)\b", question_lower):
        return ""
    if not re.search(r"\b(how many|count|number of|total)", question_lower):
        return ""

    restore_persisted_session()
    answers = []

    for filepath in core.SESSION_FILES:
        if not filepath.lower().endswith((".xlsx", ".xlsm", ".xls", ".csv")):
            continue

        rows = _spreadsheet_rows(filepath)
        if not rows:
            continue

        normalized_header = [column.strip().lower() for column in rows[0]]
        item_index = next(
            (
                index
                for index, column in enumerate(normalized_header)
                if column in {"item", "item number", "item id"}
            ),
            None
        )
        if item_index is None:
            continue

        items = {
            row[item_index].strip()
            for row in rows[1:]
            if len(row) > item_index and row[item_index].strip()
        }
        if not items:
            continue

        filename = os.path.basename(filepath)
        answers.append(
            f"There are {len(items)} distinct inventory items in {filename}. "
            f"[Source: {filename}, spreadsheet summary]"
        )

    return "\n\n".join(answers)


def build_spreadsheet_asset_filter_count_answer(question):
    """Count distinct assets whose description matches a device phrase."""

    question_lower = (question or "").lower()
    description_patterns = {
        "hard-drive": r"hard\s*drive|hardrive",
        "Unix Server": r"unix\s+server"
    }
    requested_pattern = next(
        (
            pattern
            for phrase, pattern in description_patterns.items()
            if re.search(pattern, question_lower)
        ),
        None
    )
    if requested_pattern is None:
        return ""
    if not re.search(r"\b(how many|count|number of|total)\b", question_lower):
        return ""

    restore_persisted_session()
    answers = []

    for filepath in core.SESSION_FILES:
        if not filepath.lower().endswith((".xlsx", ".xlsm", ".xls", ".csv")):
            continue

        rows = _spreadsheet_rows(filepath)
        if not rows:
            continue

        header = rows[0]
        normalized_header = [column.strip().lower() for column in header]
        description_index = next(
            (
                index
                for index, column in enumerate(normalized_header)
                if column == "description"
            ),
            None
        )
        asset_index = next(
            (
                index
                for index, column in enumerate(normalized_header)
                if column == "asset"
            ),
            None
        )
        if description_index is None or asset_index is None:
            continue

        matching_assets = {
            row[asset_index].strip()
            for row in rows[1:]
            if len(row) > max(description_index, asset_index)
            and row[asset_index].strip()
            and re.search(
                requested_pattern,
                row[description_index],
                re.IGNORECASE
            )
        }
        if not matching_assets:
            continue

        filename = os.path.basename(filepath)
        requested_label = next(
            phrase
            for phrase, pattern in description_patterns.items()
            if pattern == requested_pattern
        )
        answers.append(
            f"There are {len(matching_assets)} {requested_label} assets in {filename}. "
            f"[Source: {filename}, spreadsheet summary]"
        )

    return "\n\n".join(answers)


def build_spreadsheet_priority_answer(question):
    """Answer spreadsheet priority questions using the actual Priority column."""

    question_lower = (question or "").lower()
    if "priority" not in question_lower:
        return ""

    restore_persisted_session()
    answers = []

    for filepath in core.SESSION_FILES:
        if not filepath.lower().endswith((".xlsx", ".xlsm", ".xls", ".csv")):
            continue

        rows = _spreadsheet_rows(filepath)
        if len(rows) < 2:
            continue

        header = rows[0]
        normalized_header = [column.strip().lower() for column in header]
        priority_index = next(
            (
                index
                for index, column in enumerate(normalized_header)
                if column in {"priority", "priorities", "wopriority"}
            ),
            None
        )
        if priority_index is None:
            continue

        priorities = []
        for row in rows[1:]:
            if len(row) <= priority_index:
                continue
            value = row[priority_index].strip()
            if not value:
                continue
            match = re.search(r"-?\d+", value)
            if match:
                priorities.append(int(match.group(0)))

        if not priorities:
            continue

        filename = os.path.basename(filepath)
        requested_priorities = []
        if re.search(r"\b(highest|maximum|max)\b", question_lower):
            highest_priority = max(priorities)
            count = sum(1 for value in priorities if value == highest_priority)
            answers.append(
                f"The highest priority in {filename} is {highest_priority}, "
                f"appearing {count} time(s). [Source: {filename}, spreadsheet summary]"
            )
            continue

        for match in re.finditer(r"priority\s*(?:is|=|:)?\s*(\d+)", question_lower):
            requested_priorities.append(int(match.group(1)))

        if not requested_priorities:
            requested_priorities = [
                int(number)
                for number in re.findall(r"\b\d+\b", question_lower)
            ]

        if not requested_priorities:
            distinct_count = len(set(priorities))
            unique_values = ", ".join(str(value) for value in sorted(set(priorities)))
            answers.append(
                f"The spreadsheet contains {distinct_count} distinct priority values in {filename}: "
                f"{unique_values}. [Source: {filename}, spreadsheet summary]"
            )
            continue

        requested_priorities = sorted(set(requested_priorities))
        matching_details = []
        for requested_priority in requested_priorities:
            count = sum(1 for value in priorities if value == requested_priority)
            matching_details.append(f"priority {requested_priority}: {count}")

        if all(
            sum(1 for value in priorities if value == requested_priority) == 0
            for requested_priority in requested_priorities
        ):
            values_text = ", ".join(str(value) for value in requested_priorities)
            answers.append(
                f"There are no work orders with priority {values_text} in {filename}. "
                f"[Source: {filename}, spreadsheet summary]"
            )
            continue

        answer = (
            f"In {filename}, " + "; ".join(matching_details) +
            ". [Source: {filename}, spreadsheet summary]"
        )
        answers.append(answer)

    return "\n\n".join(answers)


def build_spreadsheet_relationship_list_answer(question):
    """List relationships by direction, and by Use With when requested."""

    question_lower = (question or "").lower()
    if not re.search(r"\b(which|what|list|show|name|identify)\b", question_lower):
        return ""
    asks_ci = bool(re.search(r"\bci\b", question_lower))
    asks_asset = bool(re.search(r"\bassets?\b", question_lower))
    asks_use_with = asks_ci or asks_asset
    if not re.search(r"\brelationships?\b", question_lower) and not (asks_ci and asks_asset):
        return ""
    asks_bidirectional = bool(re.search(r"\b(bi|bidirectional)\b", question_lower))
    asks_unidirectional = bool(re.search(r"\b(uni|unidirectional)\b", question_lower))

    restore_persisted_session()
    answers = []

    for filepath in core.SESSION_FILES:
        if not filepath.lower().endswith((".xlsx", ".xlsm", ".xls", ".csv")):
            continue

        rows = _spreadsheet_rows(filepath)
        if len(rows) < 2:
            continue

        header = [column.strip().casefold() for column in rows[0]]
        relationship_index = next(
            (index for index, column in enumerate(header) if column == "relationship"),
            None
        )
        type_index = next(
            (index for index, column in enumerate(header) if column == "type"),
            None
        )
        description_index = next(
            (index for index, column in enumerate(header) if column == "description"),
            None
        )
        use_with_index = next(
            (index for index, column in enumerate(header) if column == "use with"),
            None
        )
        if (
            relationship_index is None
            or type_index is None
            or (asks_use_with and use_with_index is None)
        ):
            continue

        grouped = {"BIDIRECTIONAL": [], "UNIDIRECTIONAL": []}
        grouped_by_use = {
            "CI": {"BIDIRECTIONAL": [], "UNIDIRECTIONAL": []},
            "ASSET": {"BIDIRECTIONAL": [], "UNIDIRECTIONAL": []}
        }
        for row in rows[1:]:
            if len(row) <= max(relationship_index, type_index):
                continue
            name = row[relationship_index].strip()
            relationship_type = re.sub(r"[^a-z]", "", row[type_index].casefold())
            if not name:
                continue
            if relationship_type == "bidirectional":
                group = "BIDIRECTIONAL"
            elif relationship_type == "unidirectional":
                group = "UNIDIRECTIONAL"
            else:
                continue

            value = name
            grouped[group].append(value)
            if use_with_index is not None and len(row) > use_with_index:
                use_with = row[use_with_index].strip().upper()
                if use_with in grouped_by_use:
                    grouped_by_use[use_with][group].append(value)

        requested_groups = []
        if asks_bidirectional:
            requested_groups.append("BIDIRECTIONAL")
        if asks_unidirectional:
            requested_groups.append("UNIDIRECTIONAL")
        if not requested_groups:
            requested_groups = ["BIDIRECTIONAL", "UNIDIRECTIONAL"]

        sections = []
        if asks_use_with:
            requested_use_with = []
            if asks_ci:
                requested_use_with.append("CI")
            if asks_asset:
                requested_use_with.append("ASSET")
            for use_with in requested_use_with:
                type_sections = []
                for group in requested_groups:
                    values = grouped_by_use[use_with][group]
                    if values:
                        label = "Bidirectional" if group == "BIDIRECTIONAL" else "Unidirectional"
                        type_sections.append(
                            f"{label} ({len(values)}): " + ", ".join(values)
                        )
                if type_sections:
                    sections.append(f"{use_with}: " + "; ".join(type_sections))
        else:
            for group in requested_groups:
                values = grouped[group]
                if values:
                    label = "Bidirectional" if group == "BIDIRECTIONAL" else "Unidirectional"
                    sections.append(f"{label} ({len(values)}): " + ", ".join(values))

        if sections:
            filename = os.path.basename(filepath)
            answers.append(
                f"Relationship types in {filename}: "
                + ". ".join(sections)
            )

    return "\n\n".join(answers)


def build_spreadsheet_status_answer(question):
    """Explain a spreadsheet status code such as APPR or WAPPR."""

    if not re.search(
        r"\b(what does|what is|meaning of|explain|define)\b",
        (question or "").lower()
    ):
        return ""

    restore_persisted_session()
    question_lower = (question or "").lower()

    for filepath in core.SESSION_FILES:
        if not filepath.lower().endswith((".xlsx", ".xlsm", ".xls", ".csv")):
            continue

        rows = _spreadsheet_rows(filepath)
        if len(rows) < 2 or "Status" not in rows[0]:
            continue

        status_index = rows[0].index("Status")
        statuses = {
            row[status_index].strip()
            for row in rows[1:]
            if len(row) > status_index and row[status_index].strip()
        }

        for status in statuses:
            if not re.search(rf"\b{re.escape(status.lower())}\b", question_lower):
                continue

            label = _status_display_name(status, question)
            filename = os.path.basename(filepath)
            return (
                f"In {filename}, status {status} means {label.title()}. "
                f"It represents work orders that are {STATUS_DESCRIPTIONS.get(label, label.lower())}. "
                f"[Source: {filename}, spreadsheet summary]"
            )

    return ""


async def save_uploaded_file(file: UploadFile, replace_session: bool = True):
    """Save, extract, and index an uploaded file for the active chat session."""

    destination = os.path.join(
        core.UPLOAD_DIR,
        file.filename
    )

    with open(destination, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    text = load_file(destination)
    if not isinstance(text, str):
        text = ""

    if not text.strip():
        print(
            "WARNING: FILE WAS NOT INDEXED BECAUSE TEXT EXTRACTION FAILED:",
            file.filename
        )
        return {
            "success": False,
            "filename": file.filename,
            "indexed": False,
            "characters": 0,
            "error": "No readable text was extracted from the file."
        }

    if replace_session:
        core.reset_chat_state()

    if destination not in core.SESSION_FILES:
        core.SESSION_FILES.append(destination)

    normalized_records = _normalized_tabular_records(destination)
    index_file(
        file.filename,
        text,
        core.ACTIVE_CONTEXT_ID,
        documents=normalized_records[0] if normalized_records else None,
        document_metadata=normalized_records[1] if normalized_records else None
    )

    print("UPLOADED:", destination)
    print("INDEXED:", file.filename)

    return {
        "success": bool(text.strip()),
        "filename": file.filename,
        "indexed": bool(text.strip()),
        "characters": len(text)
    }


def rebuild_session_index():
    """Reindex all active-session files using current normalization rules."""

    restore_persisted_session()

    for filepath in core.SESSION_FILES:

        normalized_records = _normalized_tabular_records(filepath)
        index_file(
            os.path.basename(filepath),
            load_file(filepath),
            core.ACTIVE_CONTEXT_ID,
            documents=normalized_records[0] if normalized_records else None,
            document_metadata=normalized_records[1] if normalized_records else None
        )
