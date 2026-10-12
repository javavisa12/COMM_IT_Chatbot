"""Define HTTP routes for chat, uploads, history, and knowledge inspection."""

import asyncio
import csv
import io
import json
import os
import re
import tempfile
from urllib.parse import parse_qs, quote, urlsplit

from fastapi import APIRouter, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, StreamingResponse
from openpyxl import Workbook
from pydantic import BaseModel, Field
import requests

from maximo_service import DEFAULT_SELECTS
from backend import core, chat_history
from backend.core import (
    CHAT_DIR,
    MAX_RECENT_MESSAGES,
    RECENT_MESSAGES,
    SESSION_FILES,
    reset_chat_state
)
from backend.services.chat_service import generate_response
from backend.services.file_service import (
    build_file_context,
    build_spreadsheet_column_count_answer,
    build_spreadsheet_asset_count_answer,
    build_spreadsheet_asset_filter_count_answer,
    build_spreadsheet_count_answer,
    build_spreadsheet_distinct_column_count_answer,
    build_spreadsheet_distinct_location_count_answer,
    build_spreadsheet_location_count_answer,
    build_spreadsheet_location_answer,
    build_spreadsheet_item_count_answer,
    build_spreadsheet_priority_answer,
    build_spreadsheet_relationship_list_answer,
    build_spreadsheet_status_answer,
    question_references_tabular_column,
    rebuild_session_index,
    save_uploaded_file
)
from backend.services.llm_service import llm
from backend.services.rag_service import (
    get_knowledge_explorer,
    get_knowledge_statistics
)

from mcp_client import MaximoMCPClient

mcp_client = MaximoMCPClient()

from maximo_service import (
    MaximoConfigurationError,
    MAXIMO_OBJECTS,
    get_asset,
    get_contract,
    get_item,
    get_location,
    get_object,
    get_object_count,
    get_person,
    get_purchase_order,
    get_record,
    get_tool,
    get_work_order,
    get_pm,
    get_master_pm
)

router = APIRouter()
chat_history.init_db()

# Optional hardening for the /download proxy endpoint.
#   DOWNLOAD_ALLOWED_HOSTS: comma-separated host names; when set, only those
#                           hosts may be fetched (prevents SSRF).
#   DOWNLOAD_VERIFY_SSL:    "true" to verify TLS certificates. The default
#                           keeps the previous behaviour (no verification).
DOWNLOAD_ALLOWED_HOSTS = {
    host.strip().lower()
    for host in os.getenv("DOWNLOAD_ALLOWED_HOSTS", "").split(",")
    if host.strip()
}
DOWNLOAD_VERIFY_SSL = os.getenv("DOWNLOAD_VERIFY_SSL", "false").strip().lower() in (
    "1", "true", "yes"
)

MAXIMO_DESCRIPTION_OBJECTS = {
    "asset": "MXAPIASSET",
    "tool": "MXAPITOOLINV",
    "location": "MXAPILOCATIONS",
    "contract": "MXAPICONTRACT",
    "item": "MXAPIITEM",
    "person": "MXAPIPERSON",
    "pm": "MXAPIPM",
    "master pm": "MXAPIMASTERPM"
}

MAXIMO_REFERENCE_APIS = {
    "contract": {
        "pattern": r"\bcontract\s*(?:number|no\.?|#)?\s*([A-Za-z0-9][A-Za-z0-9._-]*)",
        "identifier": "contractnum",
        "fields": (
            ("Status", "status_description", "status"),
            ("Vendor", "vendor", None),
            ("Contract type", "contracttype", None),
            ("Start date", "startdate", None),
            ("End date", "enddate", None),
            ("Total cost", "totalcost", None)
        )
    },
    "tool": {
        "pattern": r"\btools|tool?\s*(?:number|no\.?|#)?\s*([A-Za-z0-9][A-Za-z0-9._-]*)",
        "identifier": "itemnum",
        "fields": (
            ("Status", "status", "itemstatus"),
            ("Location", "location", None),
            ("Site", "siteid", None),
            ("Vendor", "vendor", None),
            ("Commodity", "commodity", None)
        )
    },
    "item": {
        "pattern": r"\bitems?\s*(?:number|no\.?|#)?\s*([A-Za-z0-9][A-Za-z0-9._-]*)",
        "identifier": "itemnum",
        "fields": (
            ("Status", "status", None),
            ("Item type", "itemtype", None),
            ("Commodity", "commodity", None)
        )
    },
    "purchase order": {
        "pattern": r"\b(?:purchase\s+orders?|pos?)\s*(?:number|no\.?|#)?\s*([A-Za-z0-9][A-Za-z0-9._-]*)",
        "identifier": "ponum",
        "fields": (
            ("Status", "status_description", "status"),
            ("Vendor", "vendor", None),
            ("Order date", "orderdate", None),
            ("Required date", "requireddate", None),
            ("Contract", "contractrefnum", None)
        )
    },
    "location": {
        "pattern": r"\blocations?\s*(?:number|no\.?|#)?\s*([A-Za-z0-9][A-Za-z0-9._-]*)",
        "identifier": "location",
        "fields": (
            ("Status", "status_description", "status"),
            ("Site", "siteid", None),
            ("Type", "type", None),
            ("Parent", "parent", None)
        )
    },
    "person": {
        "pattern": r"\b(?:people|person|employee)s?\s*(?:id|number|no\.?|#)?\s*([A-Za-z0-9][A-Za-z0-9._-]*)",
        "identifier": "personid",
        "fields": (
            ("Name", "displayname", None),
            ("Status", "status_description", "status"),
            ("Title", "title", None),
            ("Phone", "primaryphone", None),
            ("Email", "primaryemail", None)
        )
    },

     "pm": {
            "pattern": r"\bpm\s*(?:number|no\.?|#)?\s*([A-Za-z0-9][A-Za-z0-9._-]*)",
            "identifier": "pmnum",
            "fields": (
                ("Status", "status", None),
                ("Asset", "assetnum", None),
                ("Location", "location", None),
                ("Frequency", "frequency", None),
            )
        },

        "master pm": {

            "pattern": r"\bmaster\s+pm\s*"
                    r"(?:number|no\.?|#)?\s*"
                    r"([A-Za-z0-9][A-Za-z0-9._-]*)",
            "identifier": "pmnum",
            "fields": (
                ("Status", "status", None),
                ("Frequency", "frequency", None),
                ("Frequency Unit", "frequencyunit", None),
    )
},
}


MAXIMO_NON_IDENTIFIER_WORDS = {
    "a", "all", "an", "and", "are", "available", "data", "details",
    "do", "does", "for", "from", "have", "information", "is", "list",
    "me", "my", "of", "on", "please", "record", "records", "s", "show",
    "tell", "the", "this", "what", "which", "with" ,"status","active","not active"
}
MAXIMO_COLLECTION_APIS = {
    r"\b(?:purchase\s+orders?|pos?)\b": "MXAPIPO",
    r"\bcontracts?\b": "MXAPICONTRACT",
    r"\btools|tool?\b": "MXAPITOOLINV",
    r"\bitems?\b": "MXAPIITEM",
    r"\bassets?\b": "MXAPIASSET",
    r"\bwork\s*orders?\b|\bwos?\b": "MXAPIWO",
    r"\blocations?\b": "MXAPILOCATIONS",
    r"\b(?:people|persons?|employees?)\b": "MXAPIPERSON",
    r"\bmaster\s+pms?\b": "MXAPIMASTERPM",
    r"\bpms?\b": "MXAPIPM",
}
MAXIMO_CONTRACT_TYPES = {
    "purchase": "PURCHASE",

    "lease": "LEASE",
    "rental": "LEASE",
    "lease/rental": "LEASE",

    "labor rate": "LABOR",
    "labor": "LABOR",

    "master": "MASTER",

    "warranty": "WARRANTY",
}
MAXIMO_SITE_FILTERABLE_OBJECTS = {
    "MXAPIASSET",
    "MXAPIWO",
    "MXAPITOOLINV",
    "MXAPILOCATIONS",
}


MAXIMO_OBJECT_IDENTIFIERS = {
    "MXAPIASSET": "assetnum",
    "MXAPIWO": "wonum",
    "MXAPICONTRACT": "contractnum",
    "MXAPITOOLINV": "itemnum",
    "MXAPIITEM": "itemnum",
    "MXAPIPO": "ponum",
    "MXAPILOCATIONS": "location",
    "MXAPIPERSON": "personid",
    "MXAPIPM": "pmnum",
    "MXAPIMASTERPM": "pmnum",
}
MAXIMO_SELECT_ALIASES = {
    "MXAPIASSET": {
        "assetnum": ("asset number", "assetnum"),
        "location": ("location",),
        "description": ("description",),
        "status": ("status", "condition", "state"),
        "siteid": ("site", "site id", "siteid"),
        "serialnum": ("serial number", "serialnum"),
        "manufacturer": ("manufacturer",),
        "modelnum": ("model", "model number", "modelnum"),
        "parent": ("parent asset", "parent"),
        "assettype": ("asset type", "assettype"),
        "priority": ("priority",),
        "installdate": ("install date", "installation date", "installdate"),
        "warrantyexpdate": ("warranty expiry", "warranty expiration", "warrantyexpdate")
    },
    "MXAPIWO": {
        "wonum": ("work order number", "work order", "wo number", "wonum"),
        "description": ("description",),
        "status": ("status", "state"),
        "assetnum": ("asset number", "assetnum", "asset"),
        "location": ("location",),
        "siteid": ("site", "site id", "siteid"),
        "wopriority": ("priority", "wopriority"),
        "worktype": ("work type", "worktype"),
        "schedstart": ("scheduled start", "schedstart"),
        "schedfinish": ("scheduled finish", "schedfinish"),
        "reportedby": ("reported by", "reportedby"),
        "supervisor": ("supervisor",)
    },
    "MXAPITOOLINV": {
        "itemnum": ("item number", "itemnum", "tool number"),
        "description": ("description",),
        "status": ("status", "state"),
        "location": ("location",),
        "siteid": ("site", "site id", "siteid"),
        "commoditygroup": ("commodity group", "commoditygroup"),
        "issueunit": ("issue unit", "issueunit"),
        "itemtype": ("item type", "itemtype")
    },
    "MXAPIITEM": {
        "itemnum": ("item number", "itemnum"),
        "description": ("description",),
        "status": ("status", "state"),
        "itemtype": ("item type", "itemtype"),
        "commodity": ("commodity",),
        "orderunit": ("order unit", "orderunit"),
        "issueunit": ("issue unit", "issueunit"),
        "manufacturer": ("manufacturer",),
        "modelnum": ("model", "model number", "modelnum")
    },
    "MXAPICONTRACT": {
        "contractnum": ("contract number", "contractnum"),
        "description": ("description",),
        "status": ("status", "state"),
        "vendor": ("vendor",),
        "contracttype": (
            "contract type",
            "contracttype",
            "type",
            "purchase",
            "lease",
            "rental",
            "lease/rental",
            "labor rate",
            "labor",
            "master",
            "warranty"
        ),
        "contracttype_description": (
            "contract type description",
            "contract type name",
            "type description"
        ),
        "startdate": ("start date", "startdate"),
        "enddate": ("end date", "enddate"),
        "totalcost": ("total cost", "totalcost")
    },

    "MXAPIPO": {
        "ponum": ("purchase order number", "po number", "ponum"),
        "description": ("description",),
        "status": ("status", "state"),
        "vendor": ("vendor",),
        "orderdate": ("order date", "orderdate"),
        "requireddate": ("required date", "requireddate"),
        "contractrefnum": ("contract", "contract number", "contractrefnum")
    },
    "MXAPILOCATIONS": {
        "location": ("location",),
        "description": ("description",),
        "status": ("status", "state"),
        "siteid": ("site", "site id", "siteid"),
        "type": ("type", "location type"),
        "parent": ("parent", "parent location")
    },
    "MXAPIPERSON": {
        "personid": ("person id", "personid", "employee id"),
        "displayname": ("name", "display name", "displayname"),
        "status": ("status", "state"),
        "title": ("title",),
        "primaryphone": ("phone", "telephone", "primaryphone"),
        "primaryemail": ("email", "primaryemail")
    },

    "MXAPIPM": {
        "pmnum": ("pm number", "pmnum"),
        "description": ("description",),
        "status": ("status", "state"),
        "assetnum": ("asset", "asset number", "assetnum"),
        "location": ("location",),
        "siteid": ("site", "site id", "siteid"),
        "frequency": ("frequency",),
    },
    "MXAPIMASTERPM": {
        "pmnum": (
            "master pm",
            "master pm number",
            "pmnum"
        ),
        "description": ("description",),
        "status": ("status",),
        "frequency": ("frequency",),
        "frequencyunit": (
            "frequency unit",
            "frequencyunit"
        )
    },
}


def _load_maximo_object_titles():
    metadata_path = os.path.join(
        os.path.dirname(os.path.dirname(__file__)),
        "apimeta.json"
    )
    try:
        with open(metadata_path, encoding="utf-8") as metadata_file:
            metadata = json.load(metadata_file)
    except (OSError, json.JSONDecodeError):
        return {}

    titles = {}
    for entry in metadata:
        object_name = str(entry.get("osName") or "").upper()
        title = str(entry.get("title") or "").strip()
        if object_name not in MAXIMO_OBJECTS or not title:
            continue
        title = re.sub(r"^Maximo\s+(?:API|table)\s+for\s+", "", title, flags=re.IGNORECASE)
        titles[title] = object_name
    return titles


MAXIMO_OBJECT_TITLES = _load_maximo_object_titles()


class ChatRequest(BaseModel):
    message: str
    history: list[dict] = Field(default_factory=list)


class SaveChatRequest(BaseModel):
    title: str
    messages: list
    conversation_id: str | None = None


# Number of exchanges recorded since the summary was last refreshed.
_MESSAGES_SINCE_SUMMARY = 0


def update_summary():
    """Update the conversation summary after enough messages accumulate."""

    if not RECENT_MESSAGES:
        return

    summary_prompt = f"""
Existing Summary

{core.CHAT_SUMMARY}

Recent Conversation

{json.dumps(RECENT_MESSAGES, indent=2)}

Update the summary.

Keep:
- Important decisions
- Code changes
- Requirements
- User requests
- Findings

Maximum 500 words.
"""

    try:

        core.CHAT_SUMMARY = generate_response([
            {
                "role": "user",
                "content": summary_prompt
            }
        ])

        print("\nSUMMARY UPDATED\n")

    except Exception as error:
        print("SUMMARY ERROR:", error)


def maybe_update_summary():
    """Refresh the summary once every five recorded exchanges.

    Counting exchanges directly keeps the cadence correct even after
    RECENT_MESSAGES is trimmed to MAX_RECENT_MESSAGES (when its length would
    otherwise stay constant).
    """

    global _MESSAGES_SINCE_SUMMARY

    _MESSAGES_SINCE_SUMMARY += 1
    if _MESSAGES_SINCE_SUMMARY >= 5:
        _MESSAGES_SINCE_SUMMARY = 0
        update_summary()


@router.get("/")
def home():

    return FileResponse("frontend/index.html")


@router.get("/health")
def health():

    return {
        "status": "ok",
        "files_in_session": len(SESSION_FILES),
        "recent_messages": len(RECENT_MESSAGES),
        "summary_length": len(core.CHAT_SUMMARY)
    }

def target_tool_status(question):

    text = (question or "").lower()

    if "not active" in text:
        return "NOTACTIVE"

    if "inactive" in text:
        return "NOTACTIVE"

    if "active" in text:
        return "ACTIVE"

    return None



@router.get("/download")
def download_file(url: str, filename: str | None = None):
    """Fetch an external document and force a browser download to the client."""

    if not url:
        raise HTTPException(status_code=400, detail="Missing file URL")

    parts = urlsplit(url)
    if parts.scheme not in ("http", "https") or not parts.hostname:
        raise HTTPException(status_code=400, detail="Invalid file URL")
    if DOWNLOAD_ALLOWED_HOSTS and parts.hostname.lower() not in DOWNLOAD_ALLOWED_HOSTS:
        raise HTTPException(status_code=403, detail="Download host is not allowed")

    try:
        response = requests.get(
            url,
            headers={"User-Agent": "COMM_IT-AI/1.0"},
            timeout=30,
            verify=DOWNLOAD_VERIFY_SSL,
            stream=True
        )
        response.raise_for_status()
    except requests.RequestException as error:
        raise HTTPException(status_code=502, detail=f"Failed to download file: {error}") from error

    content_type = response.headers.get("Content-Type") or "application/octet-stream"
    fallback_name = (filename or url.rsplit("/", 1)[-1] or "download.pdf").strip()
    if not fallback_name.lower().endswith((".pdf", ".doc", ".docx", ".xls", ".xlsx", ".csv", ".txt")):
        fallback_name = f"{fallback_name}.pdf" if content_type.lower().startswith("application/pdf") else fallback_name

    # Keep quotes and line breaks out of the Content-Disposition header.
    fallback_name = re.sub(r'["\r\n]', "_", fallback_name)

    disposition = f'attachment; filename="{fallback_name}"'
    return StreamingResponse(
        response.iter_content(chunk_size=65536),
        media_type=content_type,
        headers={"Content-Disposition": disposition}
    )

EXPORT_COLUMNS = {

    "MXAPIASSET": [
        "assetnum",
        "description",
        "status",
        "location",
        "siteid",
        "parent"
    ],

    "MXAPIWO": [
        "wonum",
        "description",
        "status",
        "assetnum",
        "location",
        "wopriority",
        "worktype"
    ],

    "MXAPITOOLINV": [
        "itemnum",
        "description",
        "status",
        "location",
        "siteid",
        "itemtype"
    ],

    "MXAPIITEM": [
        "itemnum",
        "description",
        "status",
        "itemtype",
        "commodity"
    ],

    "MXAPICONTRACT": [
        "contractnum",
        "description",
        "status",
        "contracttype",
        "contracttype_description",
        "vendor",
        "startdate",
        "enddate"
    ],

    "MXAPIPO": [
        "ponum",
        "description",
        "status",
        "vendor",
        "orderdate"
    ],

    "MXAPILOCATIONS": [
        "location",
        "description",
        "status",
        "siteid",
        "parent"
    ],

    "MXAPIPERSON": [
        "personid",
        "displayname",
        "status",
        "title",
        "primaryemail"
    ],

    "MXAPIPM": [
        "pmnum",
        "description",
        "status",
        "assetnum",
        "location",
        "siteid",
        "frequency"
],
    "MXAPIMASTERPM": [
        "pmnum",
        "description",
        "status",
        "frequency",
        "frequencyunit"
    ]
}


def _export_select(object_name):
    """Build the comma-separated select list used by the CSV and Excel exports.

    DEFAULT_SELECTS values may be either a list or a comma-separated string,
    so both forms are handled here.
    """

    default_select = DEFAULT_SELECTS.get(object_name) or []
    if isinstance(default_select, str):
        default_select = default_select.split(",")
    return ",".join(dict.fromkeys(
        list(default_select) + list(EXPORT_COLUMNS.get(object_name) or [])
    ))


def _download_maximo_csv(object_name, filename, where=None):
    """Export Maximo records as CSV while streaming subsequent pages."""

    download_select = _export_select(object_name)

    first_page = get_object(
        object_name,
        page_size=50,
        where=where,
        page_number=1,
        select=download_select
    )

    first_records = first_page.get("member", [])
    if not first_records:
        raise HTTPException(status_code=404, detail="Maximo returned no records to download")

    def csv_pages():
        page_number = 1
        result = first_page
        records = first_records
        columns = {}

        with tempfile.TemporaryFile(mode="w+t", encoding="utf-8", newline="") as spool:
            while True:
                for record in records:
                    for field in record:
                        columns.setdefault(field, None)
                    spool.write(json.dumps(record, ensure_ascii=True))
                    spool.write("\n")

                response_info = result.get("responseInfo") or result.get("responseinfo") or {}
                next_page = response_info.get("nextPage") or response_info.get("nextpage")
                if not next_page:
                    break
                if not records:
                    raise RuntimeError("Maximo returned an empty page with a next-page link")

                next_href = next_page.get("href") if isinstance(next_page, dict) else next_page
                query = parse_qs(urlsplit(next_href).query) if isinstance(next_href, str) else {}
                next_page_number = query.get("pageno", [page_number + 1])[0]
                try:
                    next_page_number = int(next_page_number)
                except (TypeError, ValueError) as error:
                    raise RuntimeError("Maximo returned an invalid next-page number") from error
                if next_page_number <= page_number:
                    raise RuntimeError("Maximo returned a non-advancing next-page link")

                page_number = next_page_number
                result = get_object(
                    object_name,
                    page_size=50,
                    where=where,
                    page_number=page_number,
                    select=download_select
                )
                records = result.get("member", [])

            fieldnames = EXPORT_COLUMNS.get(
                object_name,
                list(columns)
            )
            output = io.StringIO(newline="")
            writer = csv.writer(output)
            writer.writerow(fieldnames)
            yield output.getvalue()

            spool.seek(0)
            for line in spool:
                record = json.loads(line)
                output = io.StringIO(newline="")
                writer = csv.writer(output)
                writer.writerow([
                    json.dumps(record.get(field), ensure_ascii=True)
                    if isinstance(record.get(field), (dict, list))
                    else record.get(field)
                    for field in fieldnames
                ])
                yield output.getvalue()

    return StreamingResponse(
        csv_pages(),
        media_type="text/csv",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )


def _download_maximo_excel(object_name, filename, where=None):
    """Fetch every Maximo page, discover all columns, and return an XLSX file."""

    download_select = _export_select(object_name)

    result = get_object(
        object_name,
        page_size=50,
        where=where,
        page_number=1,
        select=download_select
    )

    records = result.get("member", [])

    if not records:
        raise HTTPException(
            status_code=404,
            detail="Maximo returned no records to download"
        )

    columns = {}
    workbook_file = tempfile.TemporaryFile(mode="w+b")

    try:
        with tempfile.TemporaryFile(
            mode="w+t",
            encoding="utf-8"
        ) as spool:

            page_number = 1

            while True:
                for record in records:
                    for field in record:
                        columns.setdefault(field, None)

                    spool.write(json.dumps(record, ensure_ascii=True))
                    spool.write("\n")

                response_info = (
                    result.get("responseInfo")
                    or result.get("responseinfo")
                    or {}
                )

                next_page = (
                    response_info.get("nextPage")
                    or response_info.get("nextpage")
                )

                if not next_page:
                    break

                if not records:
                    raise RuntimeError(
                        "Maximo returned an empty page with a next-page link"
                    )

                next_href = (
                    next_page.get("href")
                    if isinstance(next_page, dict)
                    else next_page
                )

                query = (
                    parse_qs(urlsplit(next_href).query)
                    if isinstance(next_href, str)
                    else {}
                )

                next_page_number = query.get(
                    "pageno",
                    [page_number + 1]
                )[0]

                try:
                    next_page_number = int(next_page_number)
                except (TypeError, ValueError) as error:
                    raise RuntimeError(
                        "Maximo returned an invalid next-page number"
                    ) from error

                if next_page_number <= page_number:
                    raise RuntimeError(
                        "Maximo returned a non-advancing next-page link"
                    )

                page_number = next_page_number

                result = get_object(
                    object_name,
                    page_size=50,
                    where=where,
                    page_number=page_number,
                    select=download_select
                )

                records = result.get("member", [])

            if object_name in EXPORT_COLUMNS:
                fieldnames = EXPORT_COLUMNS[object_name]
            else:
                fieldnames = [
                    field
                    for field in columns
                    if not (
                        field.endswith("_href")
                        or field.endswith("_collectionref")
                        or field == "href"
                        or field == "_rowstamp"
                    )
                ]

            workbook = Workbook(write_only=True)

            sheet = workbook.create_sheet(
                "Tool Inventory"
                if object_name == "MXAPITOOLINV"
                else "Maximo Data"
            )

            sheet.append(fieldnames)

            # IMPORTANT:
            # spool is still open here
            spool.seek(0)

            for line in spool:
                record = json.loads(line)

                sheet.append([
                    json.dumps(record.get(field), ensure_ascii=True)
                    if isinstance(record.get(field), (dict, list))
                    else record.get(field)
                    for field in fieldnames
                ])

            workbook.save(workbook_file)

        workbook_file.seek(0)

    except Exception:
        workbook_file.close()
        raise

    def workbook_chunks():
        try:
            while chunk := workbook_file.read(65536):
                yield chunk
        finally:
            workbook_file.close()

    return StreamingResponse(
        workbook_chunks(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition":
            f'attachment; filename="{filename}"'
        }
    ) 


def download_item_list(itemnum=None):
    if itemnum is None:
        return _download_maximo_csv("MXAPIITEM", "items.csv")
    if not re.fullmatch(r"[A-Za-z0-9._-]+", itemnum):
        raise HTTPException(status_code=422, detail="Invalid item number")
    return _download_maximo_csv(
        "MXAPIITEM",
        f"item_{itemnum}.csv",
        where=f'itemnum="{itemnum}"'
    )


def download_work_orders(wonum=None):
    if wonum is None:
        return _download_maximo_csv("MXAPIWO", "work_orders.csv")
    if not re.fullmatch(r"[A-Za-z0-9._-]+", wonum):
        raise HTTPException(status_code=422, detail="Invalid work order number")
    return _download_maximo_csv(
        "MXAPIWO",
        f"work_order_{wonum}.csv",
        where=f'wonum="{wonum}"'
    )


def download_assets(assetnum=None):
    if assetnum is None:
        return _download_maximo_csv("MXAPIASSET", "assets.csv")
    if not re.fullmatch(r"[A-Za-z0-9._-]+", assetnum):
        raise HTTPException(status_code=422, detail="Invalid asset number")
    return _download_maximo_csv(
        "MXAPIASSET",
        f"asset_{assetnum}.csv",
        where=f'assetnum="{assetnum}"'
    )


def download_maximo_object(object_name, where_field=None, identifier=None):
    """Download scalar fields for any catalogued Maximo object structure."""

    object_name = (object_name or "").upper()
    if object_name not in MAXIMO_OBJECTS:
        raise HTTPException(status_code=404, detail="Unknown Maximo object structure")
    if (where_field is None) != (identifier is None):
        raise HTTPException(status_code=422, detail="Both filter field and identifier are required")

    where = None
    if where_field is not None:
        if not re.fullmatch(r"[A-Za-z][A-Za-z0-9_]*", where_field):
            raise HTTPException(status_code=422, detail="Invalid Maximo filter field")
        if not re.fullmatch(r"[A-Za-z0-9._-]+", identifier):
            raise HTTPException(status_code=422, detail="Invalid Maximo identifier")
        where = f'{where_field}="{identifier}"'

    filename = f"{object_name[5:].lower() or 'maximo'}.xlsx"
    try:
        return _download_maximo_excel(object_name, filename, where=where)
    except MaximoConfigurationError as error:
        raise HTTPException(status_code=503, detail="Maximo is not configured") from error
    except requests.HTTPError as error:
        status_code = getattr(error.response, "status_code", 502)
        raise HTTPException(
            status_code=status_code if status_code < 500 else 502,
            detail=f"Maximo API rejected the download request (HTTP {status_code})"
        ) from error
    except requests.RequestException as error:
        raise HTTPException(status_code=502, detail="Maximo API request failed") from error


LEGACY_DOWNLOAD_TOOLS = {
    "download_item_list": download_item_list,
    "download_work_orders": download_work_orders,
    "download_assets": download_assets,
}


TOOLS = {
    "download_maximo_object": download_maximo_object,
}


def download(tool_name, identifier=None):
    """Dispatch a legacy download alias or the generic catalogued-object tool."""

    tool = LEGACY_DOWNLOAD_TOOLS.get(tool_name) or TOOLS.get(tool_name)
    if tool is None:
        raise HTTPException(status_code=404, detail="Unknown download tool")
    if tool is download_maximo_object and identifier is None:
        # The generic tool requires an object name; calling it bare would
        # raise a TypeError (HTTP 500).
        raise HTTPException(status_code=422, detail="A Maximo object name is required")
    try:
        return tool(identifier) if identifier is not None else tool()
    except MaximoConfigurationError as error:
        raise HTTPException(status_code=503, detail="Maximo is not configured") from error
    except requests.HTTPError as error:
        status_code = getattr(error.response, "status_code", 502)
        raise HTTPException(
            status_code=status_code if status_code < 500 else 502,
            detail=f"Maximo API rejected the download request (HTTP {status_code})"
        ) from error
    except requests.RequestException as error:
        raise HTTPException(status_code=502, detail="Maximo API request failed") from error

def tool_status_update_request(question):
    text = (question or "").lower()

    if not re.search(
        r"\b(change|update|set|make|mark)\b",
        text
    ):
        return False

    return (
        "tool" in text
        or "tools" in text
        or "mxapitoolinv" in text
    ) and (
        "status" in text
        or "active" in text
        or "inactive" in text
        or "not active" in text
    )

@router.get("/download/tool/{tool_name}")
def download_tool(tool_name: str, identifier: str | None = None):
    return download(tool_name, identifier=identifier)


@router.get("/download/maximo/{object_name}")
def download_maximo_object_route(
    object_name: str,
    where_field: str | None = None,
    identifier: str | None = None
):
    return download_maximo_object(object_name, where_field, identifier)


@router.post("/new-chat")
def new_chat():

    reset_chat_state(clear_knowledge=False)
    return {"success": True}

@router.post("/mcp/tool/status")
async def change_tool_status(request: dict):
    return await mcp_client.call_tool(
        "change_tool_status",
        request
    )

@router.post("/upload")
async def upload_file(
    file: UploadFile = File(...),
    replace_session: bool = True
):

    return await save_uploaded_file(file, replace_session)

@router.post("/mcp/workorder/status")
async def change_status(request: dict):

    return await mcp_client.call_tool(
        "change_work_order_status",
        request
    )

@router.post("/mcp/workorder/create")
async def create_workorder(request: dict):

    return await mcp_client.call_tool(
        "create_work_order",
        request
    )

@router.post("/mcp/workorder/worklog")
async def add_worklog(request: dict):

    return await mcp_client.call_tool(
        "add_work_log",
        request
    )


@router.post("/rebuild-rag")
def rebuild_rag():

    rebuild_session_index()
    return {"success": True}


@router.get("/knowledge/stats")
def knowledge_stats():

    return get_knowledge_statistics()


@router.get("/knowledge/explorer")
def knowledge_explorer():

    return {
        "context_id": core.ACTIVE_CONTEXT_ID,
        "records": get_knowledge_explorer()
    }



def is_general_question(question):
    """Identify questions that should be answered without uploaded-file context."""

    text = (question or "").strip().lower()
    if not text:
        return False

    if question_references_tabular_column(question):
        return False

    file_keywords = [
        "file", "uploaded", "document", "pdf", "excel", "spreadsheet",
        "csv", "work order", "status", "site", "contract", "maximo",
        "annual leave", "medical leave", "sick leave", "allowance",
        "salary", "wages", "compensation", "remuneration", "benefits",
        "entitlement", "approval", "approved", "item", "description",
        "commodity group", "rotating", "kit", "condition", "asset", "location",
        "loop location", "parent", "rotating item", "m&te", "linear",
        "calibration", "scheduled start", "priority", "is dynamic",
        "dynamic job plan", "qualifications required", "this file",
        "uploaded file", "read this"
    ]

    if any(keyword in text for keyword in file_keywords):
        return False

    return bool(re.search(
        r"\b(what|who|when|where|why|how|which|define|explain|compare|tell me about)\b",
        text
    ))


def maximo_work_order_number(question):
    """Return a work-order number when the question explicitly contains one."""

    match = re.search(
        r"\b(?:work\s*order|wo)\s*(?:number|no\.?|#)?\s*([A-Za-z0-9][A-Za-z0-9._-]*)",
        question,
        re.IGNORECASE
    )
    if match and re.search(r"\d", match.group(1)):
        return match.group(1)
    return None


def maximo_asset_number(question):
    """Return an asset number when the question explicitly contains one."""

    match = re.search(
        r"\basset\s*(?:number|no\.?|#)?\s*([A-Za-z0-9][A-Za-z0-9._-]*)",
        question,
        re.IGNORECASE
    )
    if match and re.search(r"\d", match.group(1)):
        return match.group(1)
    return None

def location_details_requested(question):

    return bool(
        re.search(
            r"\b(location|locations)\b",
            question,
            re.IGNORECASE
        )
    )

def get_description_search_answer(
    object_name,
    search_term
):

    where = (
        f'description like "%{search_term}%" '
        f'or record_descriptor like "%{search_term}%"'
    )

    return get_object(
        object_name,
        page_size=20,
        where=where
    )

def maximo_location_search_term(question):

    patterns = (
        r"\btell\s+me\s+about\s+(.+?)\s+location\b",
        r"\bwhat\s+is\s+(.+?)\s+location\b",
        r"\bshow\s+(?:me\s+)?(.+?)\s+location\b",
        r"\bfind\s+(.+?)\s+location\b",
    )

    for pattern in patterns:
        match = re.search(
            pattern,
            question,
            re.IGNORECASE
        )

        if match:
            return match.group(1).strip()

    return None

def maximo_reference(question):
    """Return the Maximo object type and identifier explicitly named by a query."""

    wonum = maximo_work_order_number(question)
    if wonum:
        return "work_order", wonum

    assetnum = maximo_asset_number(question)
    if assetnum:
        return "asset", assetnum

    location_name = maximo_location_name(question)

    if location_name:
            return "location", location_name
    
    for resource, config in MAXIMO_REFERENCE_APIS.items():
        match = re.search(
            config["pattern"],
            question or "",
            re.IGNORECASE,
        )
        if not match:
            continue

        identifier = match.group(1)
        if identifier.lower() not in MAXIMO_NON_IDENTIFIER_WORDS:
            print(
                f"[REFERENCE MATCH] resource={resource} identifier={identifier}"
            )
            return resource, identifier
    return None, None


def maximo_download_identifier(question, tool_name):
    """Extract a validated Maximo identifier relevant to one export tool."""

    if tool_name == "download_work_orders":
        return maximo_work_order_number(question)
    if tool_name == "download_assets":
        return maximo_asset_number(question)
    if tool_name == "download_item_list":
        resource, identifier = maximo_reference(question)
        if resource == "item":
            return identifier
    return None


def maximo_object_query(question):
    """Parse an explicit API object and field=value lookup from chat text."""

    match = re.search(
        r"\b(MXAPI[A-Z0-9_]+)\s+([A-Za-z][A-Za-z0-9_]*)\s*=\s*([A-Za-z0-9._-]+)",
        question,
        re.IGNORECASE
    )
    if match:
        return match.group(1).upper(), match.group(2), match.group(3)
    return None


def maximo_tool_search_term(question):
    """Extract a specific tool name/keyword from a natural-language tool lookup."""

    text = (question or "").strip()
    if not text:
        return None

    patterns = (
        r"\b(?:details?|information|info)\s+(?:on|about)\s+(?:the\s+)?(.+?)\s+tool\b",
        r"\b(?:tell\s+me\s+about|what\s+is|what\s+about)\s+(?:the\s+)?(.+?)\s+tool\b",
        r"\b(?:find|show(?:\s+me)?|get|search\s+for)\s+(?:the\s+)?(.+?)\s+tool\b",
    )

    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            term = match.group(1).strip(" \"'.,:;?-\t")
            if term and term.casefold() not in {"all", "any", "the"}:
                return term

    return None

def maximo_location_name(question):

    patterns = (
        r"\btell\s+me\s+about\s+(.+?)\s+location\b",
        r"\bwhat\s+is\s+(.+?)\s+location\b",
        r"\bshow\s+(?:me\s+)?(.+?)\s+location\b",
        r"\bfind\s+(.+?)\s+location\b",
        r"\bdetails?\s+(?:on|about)\s+(.+?)\s+location\b",
    )

    for pattern in patterns:
        match = re.search(
            pattern,
            question,
            re.IGNORECASE
        )

        if match:
            return match.group(1).strip()

    return None

def maximo_collection_query(question):
    """Return a Maximo object structure for an explicit list/search request."""

    tool_search_term = maximo_tool_search_term(question)
    tool_inventory_question = re.search(
        r"\btools\b|\btool\s+(?:inventory|records|list|catalog)\b|\bhow many tools\b",
        question,
        re.IGNORECASE
    )
    collection_intent = re.search(
        r"\b(show|list|find|get|search|display|count|total|how many|number of|summarize|summary|overview|describe)\b",
        question,
        re.IGNORECASE
    )
    # Explicit API object query:
    #   "list MXAPIPM"
    #   "show MXAPIACTION"
    #   "count MXAPIASSET"
    #   "MXAPIPM records"
    explicit_object = re.search(
        r"\b(MXAPI[A-Z0-9_]+)\b",
        question or "",
        re.IGNORECASE
    )
    if explicit_object:
        explicit_name = explicit_object.group(1).upper()
        if explicit_name in MAXIMO_OBJECTS:
            return explicit_name

    if not tool_inventory_question and not collection_intent and not tool_search_term:
        return None

    for pattern, object_name in MAXIMO_COLLECTION_APIS.items():
        if re.search(pattern, question, re.IGNORECASE):
            if not maximo_reference(question)[0]:
                return object_name
    return None

def maximo_download_tool_requested(question):
    """Select a registered CSV export for an explicit download request."""

    text = (question or "").lower()
    if not re.search(r"\b(download|export|save)\b", text):
        return None

    list_request = bool(re.search(r"\b(list|all|records|as\s+csv)\b", text))
    resources = (
        ("download_work_orders", r"\bwork\s*orders?\b|\bwos?\b"),
        ("download_assets", r"\bassets?\b"),
        ("download_item_list", r"\bitems?\b")
    )
    for tool_name, pattern in resources:
        match = re.search(pattern, text)
        if match and (
            list_request
            or match.group(0).endswith("s")
            or maximo_download_identifier(text, tool_name)
        ):
            return tool_name
    return None

def maximo_contract_type(question):
    """Detect the Maximo contract type requested by the user."""

    text = (question or "").lower()

    patterns = [
        # Purchase Contract
        (r"\bpurchase\s+contracts?\b", "PURCHASE"),

        # Lease / Rental Contract
        # Maximo stores both lease and rental as LEASE
        (r"\blease\s*/\s*rental\s+contracts?\b", "LEASE"),
        (r"\blease\s+contracts?\b", "LEASE"),
        (r"\brental\s+contracts?\b", "LEASE"),
        (r"\brental\s+agreements?\b", "LEASE"),

        # Labor Contract
        # User may call it "labor rate", but Maximo stores it as LABOR
        (r"\blabor\s+rate\s+contracts?\b", "LABOR"),
        (r"\blabor\s+contracts?\b", "LABOR"),
        (r"\blabor\s+rate\b", "LABOR"),

        # Master Contract
        (r"\bmaster\s+contracts?\b", "MASTER"),

        # Warranty Contract
        (r"\bwarranty\s+contracts?\b", "WARRANTY"),
    ]

    for pattern, contract_type in patterns:
        if re.search(pattern, text, re.IGNORECASE):
            return contract_type

    return None

def maximo_contract_type_list_requested(question):
    """Detect requests asking for the available Maximo contract types."""

    text = (question or "").lower()

    patterns = (
        r"\b(?:list(?:\s+down)?|show|display|give(?:\s+me)?|what\s+are)\b"
        r".{0,30}\bcontract\s+types?\b",

        r"\bcontract\s+types?\b.{0,30}"
        r"\b(?:list|show|display|give|available|are)\b",

        r"\b(?:list|show|display)\b.{0,30}"
        r"\btypes?\b.{0,20}\bcontracts?\b",
    )

    return any(
        re.search(pattern, text, re.IGNORECASE)
        for pattern in patterns
    )

def maximo_download_object_requested(question):
    """Resolve a downloadable Maximo object from its API or natural-language name."""

    text = question or ""
    explicit_object = re.search(r"\b(MXAPI[A-Z0-9_]+)\b", text, re.IGNORECASE)
    if explicit_object:
        object_name = explicit_object.group(1).upper()
        return object_name if object_name in MAXIMO_OBJECTS else None

    for pattern, object_name in MAXIMO_COLLECTION_APIS.items():
        if re.search(pattern, text, re.IGNORECASE) and object_name in MAXIMO_OBJECTS:
            return object_name

    for title, object_name in sorted(
        MAXIMO_OBJECT_TITLES.items(),
        key=lambda entry: len(entry[0]),
        reverse=True
    ):
        if re.search(rf"(?<![A-Za-z0-9]){re.escape(title)}(?![A-Za-z0-9])", text, re.IGNORECASE):
            return object_name

    resource, _ = maximo_reference(text)
    resource_objects = {
        "work_order": "MXAPIWO",
        "asset": "MXAPIASSET",
        "contract": "MXAPICONTRACT",
        "tool": "MXAPITOOLINV",
        "item": "MXAPIITEM",
        "purchase order": "MXAPIPO",
        "location": "MXAPILOCATIONS",
        "person": "MXAPIPERSON",
        "pm": "MXAPIPM",
        "master pm": "MXAPIMASTERPM",
    }
    object_name = resource_objects.get(resource)
    if object_name in MAXIMO_OBJECTS:
        return object_name

    compact_text = re.sub(r"[^a-z0-9]", "", text.lower())
    for candidate in sorted(MAXIMO_OBJECTS, key=len, reverse=True):
        stem = candidate.removeprefix("MXAPI").lower()
        if len(stem) >= 4 and stem in compact_text:
            return candidate
    return None


def maximo_download_filter(question, object_name):
    """Build a safe identifier/type/site filter when the request names a specific export."""

    if object_name == "MXAPITOOLINV" and re.search(
        r"\btools?\s+(?:inventory|list|catalog)\b",
        question or "",
        re.IGNORECASE
    ):
        return None, None

    object_query = maximo_object_query(question)
    if object_query and object_query[0] == object_name:
        return object_query[1], object_query[2]

    resource, identifier = maximo_reference(question)
    resource_filters = {
        "work_order": ("MXAPIWO", "wonum"),
        "asset": ("MXAPIASSET", "assetnum"),
        "contract": ("MXAPICONTRACT", "contractnum"),
        "tool": ("MXAPITOOLINV", "itemnum"),
        "item": ("MXAPIITEM", "itemnum"),
        "purchase order": ("MXAPIPO", "ponum"),
        "location": ("MXAPILOCATIONS", "location"),
        "person": ("MXAPIPERSON", "personid")
    }
    resource_filter = resource_filters.get(resource)
    if resource_filter and resource_filter[0] == object_name:
        return resource_filter[1], identifier

    # Contract exports can be filtered by the requested Maximo contract type.
    if object_name == "MXAPICONTRACT":
        contract_type = maximo_contract_type(question)
        if contract_type:
            return "contracttype", contract_type

    # Collection exports filtered by site, e.g.:
    # "download/export assets for site SHENZEN".
    site_id = maximo_site_id(question)
    if site_id and object_name in MAXIMO_SITE_FILTERABLE_OBJECTS:
        return "siteid", site_id

    return None, None

def maximo_download_tool_url(question, tool_name, identifier=None):
    """Build the internal download route, including a specific record identifier."""

    url = f"/download/tool/{tool_name}"
    identifier = identifier or maximo_download_identifier(question, tool_name)
    if identifier:
        return f"{url}?identifier={quote(identifier, safe='')}"
    return url


def resolve_maximo_download_request(question, history=None):
    """Resolve a catalogued download target from the request or recent chat turns."""

    if not re.search(r"\b(download|export|save)\b", question or "", re.IGNORECASE):
        return None, None, None

    object_name = maximo_download_object_requested(question)
    if object_name:
        where_field, identifier = maximo_download_filter(question, object_name)
        return object_name, where_field, identifier

    for entry in reversed(history or []):
        if not isinstance(entry, dict):
            continue
        context = " ".join(
            str(entry.get(key) or "")
            for key in ("content", "user", "assistant")
        )
        object_name = maximo_download_object_requested(context)
        if object_name:
            where_field, identifier = maximo_download_filter(context, object_name)
            return object_name, where_field, identifier
    return None, None, None


def maximo_download_url(object_name, where_field=None, identifier=None):
    """Build the generic download route for a catalogued Maximo object."""

    url = f"/download/maximo/{quote(object_name, safe='')}"
    if where_field and identifier:
        query = f"where_field={quote(where_field, safe='')}&identifier={quote(identifier, safe='')}"
        return f"{url}?{query}"
    return url


def maximo_site_id(question):
    """Extract an explicitly supplied Maximo site ID from natural language.

    Supported examples:
      - site id = SHENZEN
      - siteid SHENZEN
      - site id: "SHENZEN"
      - site SHENZEN
      - at site SHENZEN
    """
    text = (question or "").strip()
    if not text:
        return None

    patterns = (
        r"\bsite\s*id\s*(?:is|=|:)?\s*['\"]?([A-Za-z0-9][A-Za-z0-9._-]*)['\"]?",
        r"\bsiteid\s*(?:is|=|:)?\s*['\"]?([A-Za-z0-9][A-Za-z0-9._-]*)['\"]?",
        r"\bsite\s*(?:is|=|:)\s*['\"]?([A-Za-z0-9][A-Za-z0-9._-]*)['\"]?",
        r"\b(?:at|for|from|on)\s+site\s+['\"]?([A-Za-z0-9][A-Za-z0-9._-]*)['\"]?",
        r"\b(?:at|in|for|from|on)\s+['\"]?([A-Za-z0-9][A-Za-z0-9._-]*)['\"]?\s+site\b",
    )

    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(1)

    return None


def maximo_collection_where(question, object_name):
    """Build safe OSLC filters for supported Maximo collection queries."""

    filters = []

    # ---------------------------------------------------------
    # SITE FILTER
    # ---------------------------------------------------------
    # Site filters apply only to Maximo objects that expose siteid.
    site_id = maximo_site_id(question)

    if site_id and object_name in MAXIMO_SITE_FILTERABLE_OBJECTS:
        filters.append(f'siteid="{site_id}"')

    # ---------------------------------------------------------
    # CONTRACT TYPE FILTER
    # ---------------------------------------------------------
    # Supports:
    #   purchase contracts
    #   lease contracts
    #   rental contracts
    #   labor rate contracts
    #   master contracts
    #   warranty contracts
    if object_name == "MXAPICONTRACT":
        contract_type = maximo_contract_type(question)

        if contract_type:
            filters.append(f'contracttype="{contract_type}"')

    # ---------------------------------------------------------
    # WORK ORDER PRIORITY FILTER
    # ---------------------------------------------------------
    if object_name == "MXAPIWO" and "priority" in (question or "").lower():

        question_text = (question or "").lower()

        # Prefer the numbers written right after the word "priority"
        # (e.g. "priority 2", "priority 1 and 2") so that unrelated numbers
        # such as "top 10" are not treated as priorities. Fall back to every
        # number in the question when no such phrase is found.
        scoped_match = re.search(
            r"\bpriorit(?:y|ies)\s*(?:is|are|=|:)?\s*"
            r"((?:\d+\s*(?:,|and|or|/)?\s*)+)",
            question_text
        )
        priority_source = scoped_match.group(1) if scoped_match else question_text

        priorities = [
            int(value)
            for value in re.findall(r"\b\d+\b", priority_source)
            if 1 <= int(value) <= 100
        ]

        unique_priorities = sorted(set(priorities))

        if len(unique_priorities) == 1:
            filters.append(
                f"wopriority={unique_priorities[0]}"
            )

        elif unique_priorities:
            filters.append(
                "("
                + " or ".join(
                    f"wopriority={priority}"
                    for priority in unique_priorities
                )
                + ")"
            )

    # ---------------------------------------------------------
    # TOOL SEARCH FILTER
    # ---------------------------------------------------------
    # Example:
    # "give me details on purge tool"
    #
    # Searches both item number and description while preserving
    # site/location/status filters.
  
    # TOOL SEARCH FILTER
    if object_name == "MXAPITOOLINV":
        tool_search_term = maximo_tool_search_term(question)

        if tool_search_term:
            tool_search_term = re.sub(
                r"\b(?:change|update|set|make|mark)\b", " ",
                tool_search_term, flags=re.IGNORECASE
            )
            tool_search_term = re.sub(
                r"\bstatus\b|\bto\b|\bnot\s+active\b|\binactive\b|\bactive\b",
                " ", tool_search_term, flags=re.IGNORECASE
            )
            safe_term = re.sub(r"[^A-Za-z0-9_ .-]", " ", tool_search_term).strip()
            safe_term = re.sub(r"\s+", " ", safe_term)

            if safe_term:
                filters.append(
                    f'(itemnum like "%{safe_term}%" '
                    f'or description like "%{safe_term}%")'
                )

    # ---------------------------------------------------------
    # STATUS FILTER
    # ---------------------------------------------------------
    # Examples:
    #   status=ACTIVE
    #   status: ACTIVE
    #   with status OPERATING
    #   active only
    status_match = re.search(
        r"\bstatus\s*(?:is|=|:)\s*['\"]?([A-Za-z0-9_-]+)['\"]?",
        question or "",
        re.IGNORECASE
    ) or re.search(
        r"\b(?:with|where|having)\s+status\s+['\"]?([A-Za-z0-9_-]+)['\"]?",
        question or "",
        re.IGNORECASE
    ) or re.search(
        r"\bstatus\s+(?!(?:of|in|at|on|for|with|to|from|as|report|summary|"
        r"breakdown|update|overview|change|not|is|are)\b)"
        r"['\"]?([A-Za-z0-9_-]+)['\"]?",
        question or "",
        re.IGNORECASE
    )

    if status_match:
        filters.append(
            f'status="{status_match.group(1)}"'
        )

    else:
        # Friendly shorthand for common Maximo lifecycle states.
        state_match = re.search(
            r"\b(active|inactive|operating|decommissioned)\s+only\b",
            question or "",
            re.IGNORECASE
        )

        if state_match:
            filters.append(
                f'status="{state_match.group(1)}"'
            )

    # ---------------------------------------------------------
    # LOCATION FILTER
    # ---------------------------------------------------------
    # Examples:
    #   location=KL01
    #   at location KL01
    #   in location KL01
    location_match = re.search(
        r"\blocation\s*(?:is|=|:)\s*['\"]?([A-Za-z0-9._-]+)['\"]?",
        question or "",
        re.IGNORECASE
    ) or re.search(
        r"\b(?:at|in|from)\s+location\s+['\"]?([A-Za-z0-9._-]+)['\"]?",
        question or "",
        re.IGNORECASE
    )

    if location_match:
        filters.append(
            f'location="{location_match.group(1)}"'
        )

    if object_name == "MXAPILOCATIONS":

      location_term = maximo_location_search_term(question)

      if location_term:

        filters.append(
            f'description like "%{location_term.upper()}%"'
        )
    # ---------------------------------------------------------
    # FINAL WHERE CLAUSE
    # ---------------------------------------------------------
    return " and ".join(filters) if filters else None


def maximo_collection_page_size(question, default=50):
    """Extract an explicit collection result limit without confusing priorities/IDs."""
    text = question or ""
    patterns = (
        r"\b(?:limit|show|return|give(?:\s+me)?)\s+(\d{1,3})\s+(?:records?|results?|items?|assets?|work\s*orders?|tools?|locations?|people|persons?|contracts?|purchase\s+orders?)\b",
        r"\b(?:first|top)\s+(\d{1,3})\s+(?:records?|results?|items?|assets?|work\s*orders?|tools?|locations?|people|persons?|contracts?|purchase\s+orders?)\b",
    )
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return max(1, min(int(match.group(1)), 100))
    return default


def maximo_collection_order_by(question, object_name):
    """Translate natural-language sorting into a validated Maximo field expression."""
    text = question or ""
    aliases = MAXIMO_SELECT_ALIASES.get(object_name, {})
    if not aliases:
        return None

    direction = "asc"
    if re.search(r"\b(?:descending|desc)\b", text, re.IGNORECASE):
        direction = "desc"
    elif re.search(r"\b(?:ascending|asc)\b", text, re.IGNORECASE):
        direction = "asc"

    # Prefer explicit "order by/sorted by" wording.
    for field, field_aliases in aliases.items():
        for alias in sorted(field_aliases, key=len, reverse=True):
            if re.search(
                rf"\b(?:sorted|ordered|order|sort)\s+by\s+{re.escape(alias)}\b",
                text,
                re.IGNORECASE
            ):
                return f"{field} {direction}"
    return None


def tool_inventory_summary_requested(question):
    """Detect questions asking for aggregate tool-inventory statistics."""

    return bool(re.search(
        r"\b(summary|summarize|overview|how many|count|distribution|breakdown|"
        r"concentration|largest|most common|most tools|across locations?)\b",
        question,
        re.IGNORECASE
    ))

def maximo_named_entity_search(question):

    patterns = (
        r"\btell\s+me\s+about\s+(.+?)\s+(asset|tool|location|contract|item|person|pm|master pm)\b",
        r"\bwhat\s+is\s+(.+?)\s+(asset|tool|location|contract|item|person|pm|master pm)\b",
        r"\bshow\s+(?:me\s+)?(.+?)\s+(asset|tool|location|contract|item|person|pm|master pm)\b",
        r"\bfind\s+(.+?)\s+(asset|tool|location|contract|item|person|pm|master pm)\b",
        r"\bdetails?\s+(?:on|about)\s+(.+?)\s+(asset|tool|location|contract|item|person|pm|master pm)\b",
    )

    for pattern in patterns:
        match = re.search(
            pattern,
            question,
            re.IGNORECASE
        )

        if match:
            return (
                match.group(2).lower(),
                match.group(1).strip()
            )

    return None


def maximo_collection_count_requested(question):
    """Detect requests for an exact count of Maximo collection records."""

    return bool(re.search(
        r"\b(how many|count|total|number of)\b",
        question,
        re.IGNORECASE
    ))


def maximo_document_download_requested(question):
    """Detect only genuine PM/procedure document download requests."""

    text = (question or "").lower()
    has_download_intent = bool(re.search(
        r"\b(download|attached document|attached documents|attachment|download file|download the\s+pm)\b",
        text,
        re.IGNORECASE
    ))
    has_pm_reference = bool(re.search(
        r"\b(pm|p.m.|procedure|preventive|maintenance|pump|work instruction|document)\b",
        text,
        re.IGNORECASE
    ))
    return has_download_intent and has_pm_reference


def _is_downloadable_document_url(value):
    """Return True only for file-like URLs rather than plain web pages."""

    if not isinstance(value, str):
        return False

    candidate = value.strip()
    if not re.search(r"https?://", candidate, re.IGNORECASE):
        return False

    lowered = candidate.lower()
    if re.search(r"\.(pdf|docx?|xlsx?|csv|txt|pptx?)(?:[?#].*)?$", lowered):
        return True

    if any(token in lowered for token in ("/download", "/attachment", "/docs/", "/file", "/documents/")):
        return True

    return False

def maximo_status_change_requested(question):
    return bool(re.search(
        r"\b(change|update|set|make|mark)\b.{0,60}\bstatus\b|"
        r"\bstatus\b.{0,30}\bto\b",
        question or "", re.IGNORECASE
    ))

def _collect_download_hrefs(value, seen=None):
    """Recursively collect only document-like URL values from Maximo metadata."""

    if seen is None:
        seen = set()
    hrefs = []

    if isinstance(value, dict):
        for key, nested_value in value.items():
            lower_key = str(key).lower()
            if lower_key in {
                "href", "url", "fileurl", "documenturl", "downloadurl",
                "doclink", "doclinkhref", "link"
            } and isinstance(nested_value, str) and _is_downloadable_document_url(nested_value):
                hrefs.append(nested_value)
            hrefs.extend(_collect_download_hrefs(nested_value, seen))
    elif isinstance(value, list):
        for item in value:
            hrefs.extend(_collect_download_hrefs(item, seen))
    elif isinstance(value, str) and _is_downloadable_document_url(value):
        hrefs.append(value)

    for href in hrefs:
        if href not in seen:
            seen.add(href)
    return list(seen)


def _extract_download_title(record):
    """Pick the best readable attachment title from a Maximo document record."""

    for nested_field in ("doclinks", "attachments", "documents", "documentlinks"):
        nested_value = record.get(nested_field)
        if isinstance(nested_value, list):
            for item in nested_value:
                if isinstance(item, dict):
                    for key in ("title", "name", "filename", "fileName", "docname", "documentname"):
                        value = item.get(key)
                        if isinstance(value, str) and value.strip():
                            return value.strip()
        elif isinstance(nested_value, dict):
            for key in ("title", "name", "filename", "fileName", "docname", "documentname"):
                value = nested_value.get(key)
                if isinstance(value, str) and value.strip():
                    return value.strip()

    for key in (
        "description", "record_descriptor", "docname", "documentname",
        "filename", "fileName", "title", "name", "docdesc", "docdescription"
    ):
        value = record.get(key)
        if isinstance(value, str) and value.strip():
            return value.strip()
    return "PM document"


def _build_download_url(href):
    """Build a same-origin route that downloads the file rather than opening it."""

    href = href.strip()
    filename = href.rsplit("/", 1)[-1] or "download.pdf"
    encoded_url = quote(href, safe="")
    encoded_name = quote(filename, safe="")
    return f"/download?url={encoded_url}&filename={encoded_name}"


def find_maximo_document_download(question):
    """Return a direct Maximo document download link when the user asks for one."""

    if not maximo_document_download_requested(question):
        return ""

    try:
        filters = [
            "pump",
            "maintenance",
            "procedure",
            "pm"
        ]
        for keyword in filters:
            if keyword.lower() not in (question or "").lower():
                continue
            where = " or ".join(
                f"description like '%{keyword}%'"
                for keyword in [keyword]
            )
            for object_name in ("MXAPIPM", "MXAPIMASTERPM", "MXAPIDOCINFO"):
                try:
                    result = get_object(object_name, page_size=10, where=where)
                except (MaximoConfigurationError, ValueError, requests.HTTPError, requests.RequestException):
                    continue
                for record in result.get("member", []):
                    hrefs = _collect_download_hrefs(record)
                    if hrefs:
                        return _build_download_url(hrefs[0])
            break

        for object_name in ("MXAPIPM", "MXAPIMASTERPM", "MXAPIDOCINFO"):
            try:
                result = get_object(object_name, page_size=10)
            except (MaximoConfigurationError, ValueError, requests.HTTPError, requests.RequestException):
                continue
            for record in result.get("member", []):
                hrefs = _collect_download_hrefs(record)
                if hrefs:
                    return _build_download_url(hrefs[0])
    except Exception:
        return ""

    return (
        "I found the PM record, but the Maximo API did not expose a downloadable "
        "attachment link in this response. Open the PM record in Maximo and use the "
        "attached document link there."
    )


def build_spreadsheet_answer(question):
    """Try the existing deterministic spreadsheet answer builders in priority order."""

    answer_builders = (
        build_spreadsheet_relationship_list_answer,
        build_spreadsheet_priority_answer,
        build_spreadsheet_status_answer,
        build_spreadsheet_distinct_column_count_answer,
        build_spreadsheet_distinct_location_count_answer,
        build_spreadsheet_location_count_answer,
        build_spreadsheet_asset_filter_count_answer,
        build_spreadsheet_location_answer,
        build_spreadsheet_asset_count_answer,
        build_spreadsheet_item_count_answer,
        build_spreadsheet_count_answer,
        build_spreadsheet_column_count_answer
    )
    for build_answer in answer_builders:
        answer = build_answer(question)
        if answer:
            return answer
    return ""


def clear_conversation_memory():
    """Clear the in-memory summary and recent turns."""

    global _MESSAGES_SINCE_SUMMARY

    RECENT_MESSAGES.clear()
    core.CHAT_SUMMARY = ""
    _MESSAGES_SINCE_SUMMARY = 0


def get_maximo_work_order_answer(wonum):
    """Return a readable Maximo work-order summary without invoking the LLM."""

    try:
        # Request woactivity to retrieve the associated tasks
        result = get_work_order(
            wonum, 
            select="wonum,description,status,status_description,assetnum,location,wopriority,worktype,schedstart,schedfinish,supervisor,woactivity{taskid,description,status_description,status}"
        )
    except MaximoConfigurationError:
        return "Maximo is not configured. Set MAXIMO_API_KEY in .env."
    except requests.HTTPError as error:
        status_code = getattr(error.response, "status_code", 502)
        return f"Maximo rejected the work-order request (HTTP {status_code})."
    except requests.RequestException:
        return "The Maximo API request failed."

    if not result.get("member"):
        return f"No Maximo record was found for work order {wonum}."

    record = result["member"][0]
    status_description = record.get("status_description")
    status_code = record.get("status")
    lines = [f"Work order {record.get('wonum', wonum)}"]
    if status_description:
        lines[0] += f" is {status_description}."
    elif status_code:
        lines[0] += f" has status {status_code}."
    if record.get("description"):
        lines.extend(["", f"Description: {record['description']}"])

    fields = (
        ("Site", "siteid"),
        ("Asset", "assetnum"),
        ("Location", "location"),
        ("Priority", "wopriority"),
        ("Work type", "worktype"),
        ("Scheduled start", "schedstart"),
        ("Scheduled finish", "schedfinish"),
        ("Supervisor", "supervisor")
    )
    lines.extend(
        f"{label}: {record[field]}"
        for label, field in fields
        if record.get(field) not in (None, "")
    )

    # Extract and format task details
    tasks = record.get("woactivity", [])
    if tasks:
        lines.append("\nTasks:")
        for task in tasks:
            task_id = task.get("taskid", "Unknown")
            task_desc = task.get("description", "No description")
            task_status = task.get("status_description") or task.get("status", "No status")
            lines.append(f"  - Task {task_id}: {task_desc} ({task_status})")

    return "\n".join(lines)

def get_maximo_asset_answer(assetnum):
    """Return a readable Maximo asset summary without invoking the LLM."""

    try:
        result = get_asset(assetnum)
    except MaximoConfigurationError:
        return "Maximo is not configured. Set MAXIMO_API_KEY in .env."
    except requests.HTTPError as error:
        status_code = getattr(error.response, "status_code", 502)
        return f"Maximo rejected the asset request (HTTP {status_code})."
    except requests.RequestException:
        return "The Maximo API request failed."

    if not result.get("member"):
        return f"No Maximo record was found for asset {assetnum}."

    record = result["member"][0]
    status_description = record.get("status_description")
    status_code = record.get("status")
    lines = [f"Asset {record.get('assetnum', assetnum)}"]
    if status_description:
        lines[0] += f" is {status_description}."
    elif status_code:
        lines[0] += f" has status {status_code}."
    if record.get("description"):
        lines.extend(["", f"Description: {record['description']}"])

    fields = (
        ("Site", "siteid"),
        ("Location", "location"),
        ("Parent asset", "parent"),
        ("Asset type", "assettype"),
        ("Serial number", "serialnum"),
        ("Manufacturer", "manufacturer"),
        ("Model", "modelnum")
    )
    lines.extend(
        f"{label}: {record[field]}"
        for label, field in fields
        if record.get(field) not in (None, "")
    )
    return "\n".join(lines)


def get_maximo_reference_answer(resource, identifier):
    """Fetch and format a supported Maximo record without chat or file context."""

    if resource == "work_order":
        return get_maximo_work_order_answer(identifier)
    if resource == "asset":
        return get_maximo_asset_answer(identifier)

    config = MAXIMO_REFERENCE_APIS[resource]
    try:
        resource_names = {
            "contract": "contracts",
            "tool": "tools",
            "item": "items",
            "purchase order": "purchase-orders",
            "location": "locations",
            "person": "people"
        }
        result = get_maximo_api_record(resource_names[resource], identifier)
    except MaximoConfigurationError:
        return "Maximo is not configured. Set MAXIMO_API_KEY in .env."
    except requests.HTTPError as error:
        status_code = getattr(error.response, "status_code", 502)
        return f"Maximo rejected the {resource} request (HTTP {status_code})."
    except requests.RequestException:
        return "The Maximo API request failed."

    if not result.get("member"):
        return f"No Maximo record was found for {resource} {identifier}."

    record = result["member"][0]
    identifier_value = record.get(config["identifier"], identifier)
    lines = [f"{resource.title()} {identifier_value}"]
    status_description = record.get("status_description")
    status_code = record.get("status")
    if status_description:
        lines[0] += f" is {status_description}."
    elif status_code:
        lines[0] += f" has status {status_code}."

    description = (
        record.get("description")
        or record.get("record_descriptor")
        or record.get("description_longdescription")
    )
    if description:
        lines.extend(["", f"Description: {description}"])

    for label, field, fallback in config["fields"]:
        value = record.get(field)
        if value in (None, "") and fallback:
            value = record.get(fallback)
        if value not in (None, "") and value not in (status_description, status_code):
            lines.append(f"{label}: {value}")
    return "\n".join(lines)


def get_maximo_object_answer(object_name, identifier_field, identifier):
    """Query any catalogued Maximo object and summarize its first match."""

    try:
        result = get_record(object_name, identifier_field, identifier)
    except MaximoConfigurationError:
        return "Maximo is not configured. Set MAXIMO_API_KEY in .env."
    except ValueError as error:
        return f"Invalid Maximo API query: {error}."
    except requests.HTTPError as error:
        status_code = getattr(error.response, "status_code", 502)
        return f"Maximo rejected the API query (HTTP {status_code})."
    except requests.RequestException:
        return "The Maximo API request failed."

    records = result.get("member", [])
    if not records:
        return f"No record found in {object_name} where {identifier_field} is {identifier}."

    record = records[0]
    lines = [f"{object_name} record"]
    lines.append(f"{identifier_field}: {record.get(identifier_field, identifier)}")
    description = (
        record.get("description")
        or record.get("record_descriptor")
        or record.get("description_longdescription")
    )
    if description:
        lines.append(f"Description: {description}")

    display_fields = (
        "status_description", "status", "siteid", "orgid", "location",
        "vendor", "contractnum", "ponum", "itemnum", "personid", "wonum",
        "assetnum", "startdate", "enddate", "orderdate", "requireddate"
    )
    for field in display_fields:
        value = record.get(field)
        if field != identifier_field and value not in (None, ""):
            label = field.replace("_", " ").title()
            lines.append(f"{label}: {value}")
    return "\n".join(lines)


def search_saved_history(identifier):
    """Find recent saved-chat messages that mention the exact record identifier."""

    identifier_pattern = re.compile(
        rf"(?<![A-Za-z0-9_]){re.escape(identifier)}(?![A-Za-z0-9_])",
        re.IGNORECASE
    )
    matches = []
    filenames = chat_history.list_history()[:20]
    for filename in filenames:
        for message in chat_history.load_chat(filename):
            if not isinstance(message, dict):
                continue
            content = message.get("content")
            if content is None:
                content = message.get("user") or message.get("assistant") or ""
            if identifier_pattern.search(str(content)):
                role = message.get("role", "message")
                excerpt = str(content).strip().replace("\n", " ")[:400]
                matches.append(f"{role.title()}: {excerpt}")
                if len(matches) >= 3:
                    return f"Saved chat {filename}: " + " | ".join(matches)
    if matches:
        return f"Saved chat {filenames[0]}: " + " | ".join(matches)
    return ""


def maximo_collection_request_options(question, object_name):
    """Resolve the live collection options requested in natural language."""
    target = {"kind": "collection", "object_name": object_name}
    return {
        "page_size": maximo_collection_page_size(question),
        "select": maximo_select_fields(question, target),
        "order_by": maximo_collection_order_by(question, object_name),
    }


def get_maximo_collection_answer(
    object_name,
    page_size=10,
    where=None,
    summarize_tools=False,
    exact_count=False,
    select=None,
    order_by=None
):
    """Fetch a page of records from a catalogued Maximo object structure."""

    try:
        if exact_count:
            total = get_object_count(object_name, where=where)
            resource_names = {
                "MXAPITOOLINV": "tools",
                "MXAPIWO": "work orders",
                "MXAPIASSET": "assets",
                "MXAPIPO": "purchase orders",
                "MXAPICONTRACT": "contracts",
                "MXAPIITEM": "items",
                "MXAPILOCATIONS": "locations",
                "MXAPIPERSON": "people",
                "MXAPIPM": "preventive maintenance",
                "MXAPIMASTERPM": "master preventive maintenance",
            }
            resource_name = resource_names.get(object_name, "records")
            return f"Maximo returned {total} {resource_name} matching your query."
        result = get_object(
            object_name,
            page_size=page_size,
            where=where,
            select=select,
            order_by=order_by
        )
    except MaximoConfigurationError:
        return "Maximo is not configured. Set MAXIMO_API_KEY in .env."
    except ValueError as error:
        return f"Invalid Maximo API query: {error}."
    except requests.HTTPError as error:
        status_code = getattr(error.response, "status_code", 502)
        return f"Maximo rejected the API query (HTTP {status_code})."
    except requests.RequestException:
        return "The Maximo API request failed."

    records = result.get("member", [])
    if not records:
        return f"Maximo returned no records for {object_name}."

    if object_name == "MXAPITOOLINV":
        if summarize_tools:
            return build_tool_inventory_summary(records)
        return build_tool_inventory_list(records)

    lines = [f"{object_name} returned {len(records)} record(s):"]
    for record in records:
        identity = next((
            record[field]
            for field in (
                "wonum", "assetnum", "contractnum", "ponum", "itemnum",
                "location", "personid"
            )
            if record.get(field) not in (None, "")
        ), "Record")
        description = (
            record.get("description")
            or record.get("record_descriptor")
            or record.get("description_longdescription")
        )
        status = record.get("status_description") or record.get("status")
        detail = f"{identity}: {description}" if description else str(identity)
        if status:
            detail += f" ({status})"
        lines.append(f"- {detail}")
    return "\n".join(lines)


def maximo_summary_target(question):
    """Resolve a Maximo summary/lookup into one API operation description."""

    object_query = maximo_object_query(question)
    if object_query:
        return {"kind": "record", "object_query": object_query}

    collection_object = maximo_collection_query(question)
    if collection_object:
        return {"kind": "collection", "object_name": collection_object}

    resource, identifier = maximo_reference(question)
    if resource:
        return {"kind": "reference", "resource": resource, "identifier": identifier}

    if re.search(r"\b(summary|summarize|overview|describe)\b", question, re.IGNORECASE):
        object_name = maximo_download_object_requested(question)
        if object_name:
            where_field, identifier = maximo_download_filter(question, object_name)
            if where_field and identifier:
                return {
                    "kind": "record",
                    "object_query": (object_name, where_field, identifier)
                }
            return {"kind": "collection", "object_name": object_name}
    return None


def maximo_select_fields(question, target):
    """Select only object fields mentioned by the user, plus the record key."""

    if not target:
        return None
    if target["kind"] == "record":
        object_name, identifier_field, _ = target["object_query"]
    elif target["kind"] == "reference":
        resource = target["resource"]
        resource_objects = {
            "work_order": "MXAPIWO",
            "asset": "MXAPIASSET",
            "contract": "MXAPICONTRACT",
            "tool": "MXAPITOOLINV",
            "item": "MXAPIITEM",
            "purchase order": "MXAPIPO",
            "location": "MXAPILOCATIONS",
            "person": "MXAPIPERSON"
        }
        object_name = resource_objects.get(resource)
        identifier_field = MAXIMO_OBJECT_IDENTIFIERS.get(object_name)
    else:
        object_name = target["object_name"]
        identifier_field = MAXIMO_OBJECT_IDENTIFIERS.get(object_name)

    object_aliases = MAXIMO_SELECT_ALIASES.get(object_name)
    if not object_aliases:
        return None

    requested_fields = []
    normalized_question = (question or "").casefold()
    for field, aliases in object_aliases.items():
        if any(
            re.search(
                rf"(?<![A-Za-z0-9]){re.escape(alias)}(?![A-Za-z0-9])",
                normalized_question,
                re.IGNORECASE
            )
            for alias in aliases
        ):
            requested_fields.append(field)

    if not requested_fields:
        return None
    selected_fields = list(dict.fromkeys(
        field for field in (identifier_field, *requested_fields) if field
    ))
    if selected_fields == [identifier_field]:
        return None
    return ",".join(selected_fields)


def classify_chat_intent(question, history=None):
    """Classify a chat request before choosing the API or language-model path."""

    if resolve_maximo_download_request(question, history)[0]:
        return "download"
    if maximo_document_download_requested(question):
        return "download"
    if maximo_summary_target(question):
        return "summary"
    return "general"


def get_maximo_summary_answer(question, target):
    """Fetch raw Maximo data first, then ask Qwen to answer the user's question."""

   
    select = maximo_select_fields(question, target)
    try:
        if target["kind"] == "record":
            object_name, identifier_field, identifier = target["object_query"]
            if select:
                maximo_data = get_record(
                    object_name,
                    identifier_field,
                    identifier,
                    select=select
                )
            else:
                maximo_data = get_record(object_name, identifier_field, identifier)
            object_label = object_name
        elif target["kind"] == "reference":
            resource = target["resource"]
            identifier = target["identifier"]
            reference_getters = {
                "work_order": get_work_order,
                "asset": get_asset,
                "contract": get_contract,
                "tool": get_tool,
                "item": get_item,
                "purchase order": get_purchase_order,
                "location": get_location,
                "person": get_person
            }
            if select:
                maximo_data = reference_getters[resource](identifier, select=select)
            else:
                maximo_data = reference_getters[resource](identifier)
            object_label = f"{resource} {identifier}"
        else:
            object_name = target["object_name"]
            where = maximo_collection_where(question, object_name)
            if maximo_collection_count_requested(question):
                maximo_data = {
                    "object": object_name,
                    "exact_count": get_object_count(object_name, where=where)
                }
            else:
                request_options = {"select": select} if select else {}
                maximo_data = get_object(
                    object_name,
                    page_size=50,
                    where=where,
                    **request_options
                )
            
            object_label = object_name
    except MaximoConfigurationError:
        return "Maximo is not configured. Set MAXIMO_API_KEY in .env."
    except ValueError as error:
        return f"Invalid Maximo API query: {error}."
    except requests.HTTPError as error:
        status_code = getattr(error.response, "status_code", 502)
        return f"Maximo rejected the API query (HTTP {status_code})."
    except requests.RequestException:
        return "The Maximo API request failed."

    prompt = f"""Answer the user's Maximo question using only the live API result below.
Summarize the relevant fields clearly. Do not dump the raw JSON, invent missing values,
or claim the result is exhaustive unless the API result contains an exact_count.

Question: {question}
Maximo object: {object_label}
Live Maximo API result:
{json.dumps(maximo_data, ensure_ascii=True, indent=2)}
"""
    try:
        return generate_response([{"role": "user", "content": prompt}])
    except Exception as error:
        print(f"[Intent engine] Qwen summary failed: {type(error).__name__}")
        return "Maximo data was retrieved, but Qwen could not summarize it right now."


def build_tool_inventory_summary(records):
    """Summarize returned Maximo tool records by status, site, and location."""

    if not records:
        return "Maximo returned no tool records."

    statuses = {}
    sites = {}
    locations = {}
    active_count = 0

    for record in records:
        status = str(
            record.get("status_description") or record.get("status") or "Unknown"
        ).strip()
        statuses[status] = statuses.get(status, 0) + 1
        if status.casefold() == "active":
            active_count += 1

        for field, counts in (("siteid", sites), ("location", locations)):
            value = str(record.get(field) or "").strip()
            if value:
                counts[value] = counts.get(value, 0) + 1

    coverage = []
    if locations:
        coverage.append(
            f"{len(locations)} location{'s' if len(locations) != 1 else ''}"
        )
    if sites:
        coverage.append(f"{len(sites)} site{'s' if len(sites) != 1 else ''}")
    coverage_text = f" across {' and '.join(coverage)}" if coverage else ""
    summary = [
        f"Maximo returned {len(records)} tool records{coverage_text} in this "
        "result (the request limit is 50 records)."
    ]

    if active_count > len(records) / 2:
        percentage = round(active_count * 100 / len(records))
        summary.append(
            f"Most tools are active: {active_count} of {len(records)} "
            f"({percentage}%)."
        )
    elif statuses:
        status_breakdown = ", ".join(
            f"{status}: {count}"
            for status, count in sorted(
                statuses.items(), key=lambda item: (-item[1], item[0].casefold())
            )
        )
        summary.append(f"Status breakdown: {status_breakdown}.")

    for label, counts in (("Site", sites), ("Location", locations)):
        if not counts:
            continue
        highest_count = max(counts.values())
        most_common = sorted(
            name for name, count in counts.items() if count == highest_count
        )
        if len(most_common) == 1:
            summary.append(
                f"Largest concentration by {label.lower()}: {most_common[0]} "
                f"({highest_count} tools)."
            )
        else:
            tied_names = " and ".join(most_common)
            summary.append(
                f"Largest concentration by {label.lower()} is tied between "
                f"{tied_names} ({highest_count} tools each)."
            )

    return " ".join(summary)


def build_tool_inventory_list(records):
    """List the returned tools without adding inventory-wide statistics."""

    lines = [
        f"Tools returned by Maximo: {len(records)} "
        "(request limit: 50 records)."
    ]
    for record in records:
        identifier = record.get("itemnum") or record.get("toolnum") or "Tool"
        description = (
            record.get("description")
            or record.get("record_descriptor")
            or record.get("description_longdescription")
        )
        details = []
        status = record.get("status_description") or record.get("status")
        if status:
            details.append(str(status))
        if record.get("siteid"):
            details.append(f"Site {record['siteid']}")
        if record.get("location"):
            details.append(f"Location {record['location']}")

        line = f"- {identifier}"
        if description:
            line += f": {description}"
        if details:
            line += f" ({'; '.join(details)})"
        lines.append(line)
    return "\n".join(lines)


def get_maximo_chat_answer(
    question,
    reference=None,
    object_query=None,
    collection_object=None
):
    """Search history and uploaded context before fetching current Maximo data."""

    # A direct collection lookup is only possible when a collection object was
    # resolved; object_query/reference lookups fall through to the 3-step path.
    if collection_object and (
        collection_object == "MXAPITOOLINV"
        or maximo_collection_count_requested(question)
    ):
        print("[Chat search] Tool inventory request: using Maximo records directly")
        options = maximo_collection_request_options(question, collection_object)
        return get_maximo_collection_answer(
            collection_object,
            where=maximo_collection_where(question, collection_object),
            summarize_tools=(
                collection_object == "MXAPITOOLINV"
                and tool_inventory_summary_requested(question)
            ),
            exact_count=maximo_collection_count_requested(question),
            **options
        )

    identifier = (
        object_query[2]
        if object_query
        else reference[1]
        if reference
        else question
    )
    print(f"[Chat search] Step 1/3: checking saved chat history for {identifier}")
    try:
        history_context = search_saved_history(identifier)
    except Exception as error:
        history_context = ""
        print(f"[Chat search] Saved-history search failed: {type(error).__name__}")
    print(
        "[Chat search] Saved-chat matches found"
        if history_context else "[Chat search] No matching saved-chat history"
    )

    print(f"[Chat search] Step 2/3: checking uploaded documents for: {question}")
    try:
        document_context = build_file_context(question)
    except Exception as error:
        document_context = ""
        print(f"[Chat search] Uploaded-document search failed: {type(error).__name__}")

    if document_context:
        print("[Chat search] Uploaded-document context found")
    else:
        print("[Chat search] No relevant uploaded-document context found")

    print("[Chat search] Step 3/3: querying Maximo")
    if object_query:
        api_answer = get_maximo_object_answer(*object_query)
    elif collection_object:
        collection_where = maximo_collection_where(question, collection_object)
        options = maximo_collection_request_options(question, collection_object)
        api_answer = get_maximo_collection_answer(
            collection_object,
            where=collection_where,
            summarize_tools=(
                collection_object == "MXAPITOOLINV"
                and tool_inventory_summary_requested(question)
            ),
            **options
        )
    else:
        api_answer = get_maximo_reference_answer(*reference)

    source_context = "\n\n".join(
        section
        for section in (history_context, document_context, api_answer)
        if section
    )
    final_prompt = f"""
Answer the user's question using the context below. Give one direct, polished
chatbot answer. Do not list the sources separately or describe the search steps.
Prefer saved-chat and uploaded-document context when it directly answers the
question. Use the live Maximo result only for explicit live-system lookups or
when the file/chat context is absent or unrelated. Do not let Maximo override
direct document evidence when the uploaded files answer the question precisely.
Do not invent details that are absent from the context.

Question:
{question}

Context gathered from saved conversations, uploaded documents, and the live
Maximo API:
{source_context}
"""

    try:
        return generate_response([{"role": "user", "content": final_prompt}])
    except Exception as error:
        print(
            "[Chat search] Final answer generation failed: "
            f"{type(error).__name__}"
        )
        return api_answer


def get_maximo_api_record(resource, identifier):
    """Fetch a supported Maximo object structure for a direct API endpoint."""

    if resource == "work-orders":
        return get_work_order(identifier)
    if resource == "assets":
        return get_asset(identifier)
    resource_apis = {
        "contracts": "contract",
        "tools": "tool",
        "items": "item",
        "purchase-orders": "purchase order",
        "locations": "location",
        "people": "person",
        "pms": "pm",
        "master-pms": "master pm",
    }
    key = resource_apis.get(resource)
    if not key:
        raise ValueError("Unsupported Maximo resource")
    api_getters = {
        "contract": get_contract,
        "tool": get_tool,
        "item": get_item,
        "purchase order": get_purchase_order,
        "location": get_location,
        "person": get_person,
        "pm": get_pm,
        "master pm": get_master_pm
    }
    return api_getters[key](identifier)


def build_prompt(question):
    """Create the LLM prompt containing conversation and uploaded-file context."""

    resource, identifier = maximo_reference(question)
    if resource or is_general_question(question):
        file_context = ""
    else:
        file_context = build_file_context(question)

    maximo_context = ""
    if resource:
        try:
            maximo_context = json.dumps(
                get_maximo_api_record(
                    {
                        "work_order": "work-orders",
                        "asset": "assets",
                        "contract": "contracts",
                        "tool": "tools",
                        "item": "items",
                        "purchase order": "purchase-orders",
                        "location": "locations",
                        "person": "people"
                    }[resource],
                    identifier
                ),
                indent=2
            )
        except requests.RequestException:
            maximo_context = "The Maximo API lookup failed. Do not guess the work-order data."
        except MaximoConfigurationError:
            maximo_context = "Maximo is not configured. Do not guess the work-order data."

    is_maximo_lookup = bool(resource)
    conversation_summary = "" if is_maximo_lookup else core.CHAT_SUMMARY
    recent_messages = [] if is_maximo_lookup else RECENT_MESSAGES

    file_context_label = (
        file_context
        if file_context
        else "No directly relevant uploaded-file context was found. This section is informational only; ignore it when the question is general and unrelated to the uploaded files."
    )

    return f"""
You are a Senior Enterprise Software Engineer.

Conversation Summary:

{conversation_summary}

Recent Messages:

{json.dumps(recent_messages, indent=2)}

Relevant Context Retrieved From Files:

{file_context_label}

Live Maximo API Context:

{maximo_context or "No specific work order was requested from the Maximo API."}

Current Question:

{question}

Instructions:

- Use the retrieved file context only when it is directly relevant to the question.
- If there is no relevant file context, ignore this section and answer using general knowledge.
- Do not say that the uploaded files contain no information for general questions that are unrelated to the uploaded files.
- Do not say that you have no knowledge or that the answer is unavailable just because it is not in the uploaded files.
- For count, total, or "how many" questions about a spreadsheet, count the complete retrieved spreadsheet data. Do not estimate from a partial sample.
- Mention filenames when the answer uses uploaded-file context.
- Cite file context using the exact format [Source: filename, chunk N].
- Only cite sources that appear in the retrieved context.
- Do not invent details about uploaded files.
- Treat Maximo API context as the source of truth for explicitly requested work orders.
- If the Maximo API lookup failed, say that the work order could not be retrieved; do not invent its details.
- If the user explicitly asks what an uploaded file says and the answer is not present in the relevant file context, say that it was not found in the uploaded files.
"""


@router.get("/maximo/work-orders/{wonum}")
async def maximo_work_order(wonum: str):
    """Return a work order fetched from the configured Maximo API."""

    try:
        result = await mcp_client.call_tool(
        "search_work_order",
        {
            "wonum": wonum
        }
               )
    except ValueError as error:
        raise HTTPException(status_code=422, detail=str(error)) from error
    except MaximoConfigurationError as error:
        raise HTTPException(status_code=503, detail=str(error)) from error
    except requests.HTTPError as error:
        status_code = getattr(error.response, "status_code", 502)
        raise HTTPException(
            status_code=status_code if status_code < 500 else 502,
            detail=f"Maximo API rejected the request (HTTP {status_code})"
        ) from error
    except requests.RequestException as error:
        raise HTTPException(
            status_code=502,
            detail="Maximo API request failed"
        ) from error


@router.get("/maximo/objects/{object_name}")
def maximo_object(
    object_name: str,
    identifier_field: str | None = None,
    identifier: str | None = None,
    page_size: int = 5,
    page_number: int | None = None,
    where: str | None = None,
    select: str | None = None,
    order_by: str | None = None
):
    """Query any allow-listed Maximo object structure with OSLC options."""

    if (identifier_field is None) != (identifier is None):
        raise HTTPException(
            status_code=422,
            detail="Provide both identifier_field and identifier, or neither"
        )
    try:
        if identifier_field is not None:
            if where or page_number is not None or order_by is not None:
                raise HTTPException(
                    status_code=422,
                    detail="where/page_number/order_by are only supported for collection queries"
                )
            return get_record(
                object_name,
                identifier_field,
                identifier,
                select=select
            )
        return get_object(
            object_name,
            page_size=page_size,
            page_number=page_number,
            where=where,
            select=select,
            order_by=order_by
        )
    except ValueError as error:
        raise HTTPException(status_code=422, detail=str(error)) from error
    except MaximoConfigurationError as error:
        raise HTTPException(status_code=503, detail=str(error)) from error
    except requests.HTTPError as error:
        status_code = getattr(error.response, "status_code", 502)
        raise HTTPException(
            status_code=status_code if status_code < 500 else 502,
            detail=f"Maximo API rejected the request (HTTP {status_code})"
        ) from error
    except requests.RequestException as error:
        raise HTTPException(
            status_code=502,
            detail="Maximo API request failed"
        ) from error


@router.get("/maximo/{resource}/{identifier}")
def maximo_record(resource: str, identifier: str):
    """Return one record from a supported Maximo object structure."""

    try:
        return get_maximo_api_record(resource, identifier)
    except ValueError as error:
        raise HTTPException(status_code=422, detail=str(error)) from error
    except MaximoConfigurationError as error:
        raise HTTPException(status_code=503, detail=str(error)) from error
    except requests.HTTPError as error:
        status_code = getattr(error.response, "status_code", 502)
        raise HTTPException(
            status_code=status_code if status_code < 500 else 502,
            detail=f"Maximo API rejected the request (HTTP {status_code})"
        ) from error
    except requests.RequestException as error:
        raise HTTPException(
            status_code=502,
            detail="Maximo API request failed"
        ) from error


def get_maximo_contract_types_answer():
    """Return the distinct contract types currently exposed by Maximo."""

    contract_types = {}

    try:
        page_number = 1
        while True:
            result = get_object(
                "MXAPICONTRACT",
                page_size=100,
                page_number=page_number,
                select="contracttype,contracttype_description"
            )
            records = result.get("member", [])
            if not records:
                break

            for record in records:
                code = str(record.get("contracttype") or "").strip()
                description = str(
                    record.get("contracttype_description") or ""
                ).strip()
                if code:
                    contract_types[code] = description or code

            response_info = (
                result.get("responseInfo")
                or result.get("responseinfo")
                or {}
            )
            next_page = (
                response_info.get("nextPage")
                or response_info.get("nextpage")
            )
            if not next_page:
                break

            next_href = (
                next_page.get("href")
                if isinstance(next_page, dict)
                else next_page
            )
            query = (
                parse_qs(urlsplit(next_href).query)
                if isinstance(next_href, str)
                else {}
            )
            next_page_number = query.get(
                "pageno",
                [page_number + 1]
            )[0]
            try:
                next_page_number = int(next_page_number)
            except (TypeError, ValueError) as error:
                raise RuntimeError(
                    "Maximo returned an invalid next-page number"
                ) from error

            if next_page_number <= page_number:
                raise RuntimeError(
                    "Maximo returned a non-advancing next-page link"
                )
            page_number = next_page_number

            if page_number > 1000:
                raise RuntimeError(
                    "Maximo contract type pagination exceeded the safety limit"
                )

    except MaximoConfigurationError:
        return "Maximo is not configured. Set MAXIMO_API_KEY in .env."
    except ValueError as error:
        return f"Invalid Maximo API query: {error}."
    except requests.HTTPError as error:
        status_code = getattr(error.response, "status_code", 502)
        return f"Maximo rejected the API query (HTTP {status_code})."
    except requests.RequestException:
        return "The Maximo API request failed."
    except RuntimeError as error:
        return f"Unable to retrieve all Maximo contract types: {error}."

    if not contract_types:
        return "Maximo returned no contract types."

    lines = ["Available Maximo contract types:"]
    for code, description in sorted(
        contract_types.items(),
        key=lambda item: (item[1].casefold(), item[0].casefold())
    ):
        if description.casefold() == code.casefold():
            lines.append(f"- {code}")
        else:
            lines.append(f"- {code}: {description}")

    return "\n".join(lines)


@router.post("/chat")
def chat(request: ChatRequest):
    """Return a complete answer, using deterministic spreadsheet logic first."""

    print(f"\nQUESTION:\n{request.message}\n")
    if tool_status_update_request(request.message):

     payload = {
        "status": "NOTACTIVE"
    }

    try:
        result = asyncio.run(
            mcp_client.call_tool(
                "change_tool_status",
                payload
            )
        )

        return {
            "response": str(result)
        }

    except Exception as e:
        return {
            "response": f"Tool status update failed: {e}"
        }
    print(f"FILES IN SESSION: {len(SESSION_FILES)}")

    object_name, where_field, identifier = resolve_maximo_download_request(
        request.message,
        request.history
    )
    if object_name:
        answer = maximo_download_url(
            object_name,
            where_field,
            identifier
        )
        return {"response": answer}

    direct_download = find_maximo_document_download(request.message)
    if direct_download:
        answer = direct_download
        print("\nMAXIMO DOWNLOAD ANSWER:\n")
        print(answer[:1000])
        RECENT_MESSAGES.append({
            "user": request.message,
            "assistant": answer
        })
        del RECENT_MESSAGES[:-MAX_RECENT_MESSAGES]
        maybe_update_summary()
        return {"response": answer}

    if re.search(r"\b(download|export|save)\b", request.message, re.IGNORECASE):
        return {"response": "Which Maximo object or record should I download?"}

    if maximo_contract_type_list_requested(request.message):
        clear_conversation_memory()
        answer = get_maximo_contract_types_answer()
        return {"response": answer}

    # Resolve explicit collection requests before the generic summary/LLM path.
    # This is critical for requests such as:
    #   "give assets based on site id = SHENZEN"
    #   "show assets for site SHENZEN"
    #   "how many assets at site SHENZEN"
    # An explicit site filter must never be replaced by spreadsheet/RAG context
    # such as an unrelated site (for example, Bedford).
    collection_object = maximo_collection_query(request.message)
    site_id = maximo_site_id(request.message)
    if collection_object and site_id:
        clear_conversation_memory()
        options = maximo_collection_request_options(
            request.message, collection_object
        )
        answer = get_maximo_collection_answer(
            collection_object,
            where=maximo_collection_where(request.message, collection_object),
            summarize_tools=(
                collection_object == "MXAPITOOLINV"
                and tool_inventory_summary_requested(request.message)
            ),
            exact_count=maximo_collection_count_requested(request.message),
            **options
        )
        print("\nFILTERED MAXIMO ANSWER:\n")
        print(answer[:1000])
        return {"response": answer}

    # Specific natural-language tool lookups must go directly to Maximo.
    # Example: "give me details on purge tool".
    tool_search_term = maximo_tool_search_term(request.message)
    if tool_search_term:
        clear_conversation_memory()
        options = maximo_collection_request_options(
            request.message, "MXAPITOOLINV"
        )
        answer = get_maximo_collection_answer(
            "MXAPITOOLINV",
            where=maximo_collection_where(request.message, "MXAPITOOLINV"),
            summarize_tools=False,
            exact_count=False,
            **options
        )
        print("\nSPECIFIC MAXIMO TOOL ANSWER:\n")
        print(answer[:1000])
        return {"response": answer}

    intent = classify_chat_intent(request.message, request.history)
    if intent == "summary":
        clear_conversation_memory()
        target = maximo_summary_target(request.message)
        answer = get_maximo_summary_answer(request.message, target)
        return {"response": answer}

    if (
        collection_object == "MXAPITOOLINV"
        or (
            collection_object
            and maximo_collection_count_requested(request.message)
        )
    ):
        clear_conversation_memory()
        answer = get_maximo_chat_answer(
            request.message,
            collection_object=collection_object
        )
        return {"response": answer}

    spreadsheet_answer = build_spreadsheet_answer(request.message)
    if spreadsheet_answer:
        answer = spreadsheet_answer
        print("\nANSWER:\n")
        print(answer[:1000])
        RECENT_MESSAGES.append({
            "user": request.message,
            "assistant": answer
        })
        del RECENT_MESSAGES[:-MAX_RECENT_MESSAGES]
        maybe_update_summary()
        return {"response": answer}

    object_query = maximo_object_query(request.message)
    if object_query:
        clear_conversation_memory()
        answer = get_maximo_chat_answer(
            request.message,
            object_query=object_query
        )
        print("\nMAXIMO ANSWER:\n")
        print(answer[:1000])
        return {"response": answer}

    collection_object = maximo_collection_query(request.message)
    if collection_object:
        clear_conversation_memory()
        answer = get_maximo_chat_answer(
            request.message,
            collection_object=collection_object
        )
        print("\nMAXIMO ANSWER:\n")
        print(answer[:1000])
        return {"response": answer}

    resource, identifier = maximo_reference(request.message)
    if resource:
        clear_conversation_memory()
        answer = get_maximo_chat_answer(
            request.message,
            reference=(resource, identifier)
        )
        print("\nMAXIMO ANSWER:\n")
        print(answer[:1000])
        return {"response": answer}

    answer = generate_response([
        {
            "role": "user",
            "content": build_prompt(request.message)
        }
    ])

    print("\nANSWER:\n")
    print(answer[:1000])

    RECENT_MESSAGES.append({
        "user": request.message,
        "assistant": answer
    })
    del RECENT_MESSAGES[:-MAX_RECENT_MESSAGES]

    maybe_update_summary()

    return {"response": answer}


@router.post("/chat/stream")
def chat_stream(request: ChatRequest):
    """Stream a deterministic spreadsheet answer or the local LLM response."""

    object_name, where_field, identifier = resolve_maximo_download_request(
        request.message,
        request.history
    )
    if object_name:
        answer = maximo_download_url(
            object_name,
            where_field,
            identifier
        )
        return StreamingResponse(iter([answer]), media_type="text/plain")

    direct_download = find_maximo_document_download(request.message)
    if direct_download:
        answer = direct_download
        RECENT_MESSAGES.append({
            "user": request.message,
            "assistant": answer
        })
        del RECENT_MESSAGES[:-MAX_RECENT_MESSAGES]
        return StreamingResponse(iter([answer]), media_type="text/plain")

    if re.search(r"\b(download|export|save)\b", request.message, re.IGNORECASE):
        return StreamingResponse(
            iter(["Which Maximo object or record should I download?"]),
            media_type="text/plain"
        )

    if maximo_contract_type_list_requested(request.message):
        clear_conversation_memory()
        answer = get_maximo_contract_types_answer()
        return StreamingResponse(iter([answer]), media_type="text/plain")

    # Explicit site-filtered collection requests must bypass the generic
    # summary/LLM path for the same reason as /chat.
    collection_object = maximo_collection_query(request.message)
    site_id = maximo_site_id(request.message)
    if collection_object and site_id:
        clear_conversation_memory()
        options = maximo_collection_request_options(
            request.message, collection_object
        )
        answer = get_maximo_collection_answer(
            collection_object,
            where=maximo_collection_where(request.message, collection_object),
            summarize_tools=(
                collection_object == "MXAPITOOLINV"
                and tool_inventory_summary_requested(request.message)
            ),
            exact_count=maximo_collection_count_requested(request.message),
            **options
        )
        return StreamingResponse(iter([answer]), media_type="text/plain")

    intent = classify_chat_intent(request.message, request.history)
    if intent == "summary":
        clear_conversation_memory()
        target = maximo_summary_target(request.message)
        answer = get_maximo_summary_answer(request.message, target)
        return StreamingResponse(iter([answer]), media_type="text/plain")

    if (
        collection_object == "MXAPITOOLINV"
        or (
            collection_object
            and maximo_collection_count_requested(request.message)
        )
    ):
        clear_conversation_memory()
        answer = get_maximo_chat_answer(
            request.message,
            collection_object=collection_object
        )
        return StreamingResponse(iter([answer]), media_type="text/plain")

    exact_answer = build_spreadsheet_answer(request.message)
    if exact_answer:
        answer = exact_answer
        RECENT_MESSAGES.append({
            "user": request.message,
            "assistant": answer
        })
        del RECENT_MESSAGES[:-MAX_RECENT_MESSAGES]
        return StreamingResponse(iter([answer]), media_type="text/plain")

    object_query = maximo_object_query(request.message)
    if object_query:
        clear_conversation_memory()
        answer = get_maximo_chat_answer(
            request.message,
            object_query=object_query
        )
        return StreamingResponse(iter([answer]), media_type="text/plain")

    resource, identifier = maximo_reference(request.message)
    collection_object = maximo_collection_query(request.message)
    if collection_object:
        clear_conversation_memory()
        answer = get_maximo_chat_answer(
            request.message,
            collection_object=collection_object
        )
        return StreamingResponse(iter([answer]), media_type="text/plain")

    if resource:
        clear_conversation_memory()
        answer = get_maximo_chat_answer(
            request.message,
            reference=(resource, identifier)
        )
        return StreamingResponse(
            iter([answer]),
            media_type="text/plain"
        )

    prompt = build_prompt(request.message)

    def generate_chunks():

        chunks = []

        try:
            for chunk in llm.create_chat_completion_stream([
                {
                    "role": "user",
                    "content": prompt
                }
            ]):
                chunks.append(chunk)
                yield chunk
        except requests.RequestException as error:
            print(
                "[Chat stream] Local model request failed: "
                f"{type(error).__name__}"
            )
            if not chunks:
                yield (
                    "The local AI model is unavailable. "
                    "Start the model server on port 8001 and try again."
                )
            return

        answer = "".join(chunks)
        RECENT_MESSAGES.append({
            "user": request.message,
            "assistant": answer
        })
        del RECENT_MESSAGES[:-MAX_RECENT_MESSAGES]

        maybe_update_summary()

    return StreamingResponse(
        generate_chunks(),
        media_type="text/plain"
    )

def target_tool_status(question):

    text = (question or "").lower()

    if "not active" in text:
        return "NOTACTIVE"

    if "inactive" in text:
        return "NOTACTIVE"

    if "active" in text:
        return "ACTIVE"

    return None



@router.post("/save-chat")
def save_chat(request: SaveChatRequest):

    filename = chat_history.save_chat(
        request.title,
        request.messages,
        request.conversation_id,
    )
    return {"success": True, "filename": filename}


@router.get("/history")
def history():

    return chat_history.list_history()


@router.get("/history/{filename}")
def load_chat(filename: str):

    return chat_history.load_chat(filename)


@router.delete("/history/{filename}")
def delete_chat(filename: str):

    return {"success": chat_history.delete_chat(filename)}