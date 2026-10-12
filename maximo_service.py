import os
import re
import time
from pathlib import Path
from urllib.parse import parse_qs, urlsplit

import requests
import urllib3
from dotenv import load_dotenv


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

urllib3.disable_warnings()

BASE_URL = "https://web.manage.demo.apps.mas9.eam-commit.com/maximo"

PROJECT_DIR = Path(__file__).resolve().parent
MAXIMO_OBJECTS_FILE = PROJECT_DIR / "apis.txt"

load_dotenv(PROJECT_DIR / ".env")


MAXIMO_OBJECTS = {
    name.strip()
    for name in MAXIMO_OBJECTS_FILE.read_text(encoding="utf-8").splitlines()
    if re.fullmatch(r"MXAPI[A-Z0-9_]+", name.strip())
}

# Built-in API catalog.  apis.txt can still contain additional objects, but
# these names are always accepted by the generic Maximo query path.
MAXIMO_OBJECTS.update(['MXAPIACTION', 'MXAPIADDRESS', 'MXAPIADMIN', 'MXAPIAHASSETAPMVAL', 'MXAPIAHBLTHRESHOLD', 'MXAPIAHCAPAD', 'MXAPIAHCLASSIFICATN', 'MXAPIAHCONTAINERCFG', 'MXAPIAHDRIVERLIB', 'MXAPIAHFACTORLIB', 'MXAPIAHHISTORY', 'MXAPIAHLOCAPMVAL', 'MXAPIAHMATRIXCFG', 'MXAPIAHMEASUREPOINT', 'MXAPIAHMETHODOLOGY', 'MXAPIAHRRPLAN', 'MXAPIAHTHRESHOLD', 'MXAPIAHWODETAIL', 'MXAPIAIDOMAINRULE', 'MXAPIAIOASSET', 'MXAPIAIOASSETSCORE', 'MXAPIAIOPROJECT', 'MXAPIAIOSCHEDASSET', 'MXAPIAIOSCHEDULE', 'MXAPIAIOSTRATEGY', 'MXAPIAIOTEMPLATE', 'MXAPIALERT', 'MXAPIALLPRODUCTS', 'MXAPIALNDOMAIN', 'MXAPIAMCREW', 'MXAPIAMCREWT', 'MXAPIANALYTICDATASET', 'MXAPIAPIKEY', 'MXAPIAPMOBJECTATTR', 'MXAPIAPMVALUE', 'MXAPIAPMVALUEATTR', 'MXAPIAPMVALUEMETHOD', 'MXAPIAPPDOCTYPE', 'MXAPIAPPPROPCFG', 'MXAPIASSET', 'MXAPIASSETATTRIBUTE', 'MXAPIASSETINT', 'MXAPIASSETMETER', 'MXAPIASSETRRPLAN', 'MXAPIASSETSPARE', 'MXAPIASSETSTATUS', 'MXAPIASSIGNMENT', 'MXAPIATTRFORMULA', 'MXAPIAUTHCI', 'MXAPIAUTOSCRIPT', 'MXAPIBILLTOSHIPTO', 'MXAPIBOOKMARK', 'MXAPIBULLETINBOARD', 'MXAPICAASSETWO', 'MXAPICALOCATIONWO', 'MXAPICLASSSTRUCTURE', 'MXAPICNTBOOK', 'MXAPICNTBOOKLINE', 'MXAPICOA', 'MXAPICOLLECTAUTH', 'MXAPICOLLECTION', 'MXAPICOMMODITY', 'MXAPICOMPONENT', 'MXAPICONTRACT', 'MXAPICONVERSION', 'MXAPICRAFT', 'MXAPICRONTASKDEF', 'MXAPICTEMPLATE', 'MXAPIDATAPROFILE', 'MXAPIDATARESTRICTION', 'MXAPIDOCINFO', 'MXAPIDOCTYPE', 'MXAPIDOCTYPES', 'MXAPIDOMAIN', 'MXAPIEMPACT', 'MXAPIENDPOINT', 'MXAPIEXPGROUP', 'MXAPIEXPGROUPCHILD', 'MXAPIFAILURECODE', 'MXAPIFAILURELIST', 'MXAPIFILEIMPORTQUEUE', 'MXAPIGLCOMP', 'MXAPIGLCONFIGURE', 'MXAPIGLTXN', 'MXAPIGROUP', 'MXAPIHAZARD', 'MXAPIINCIDENT', 'MXAPIINSIGHT', 'MXAPIINSIGHTFACTOR', 'MXAPIINSPACTIONS', 'MXAPIINSPECTIONRES', 'MXAPIINSPFDOCLINKS', 'MXAPIINSPFIELD', 'MXAPIINSPFORM', 'MXAPIINSPFORMSTATUS', 'MXAPIINSPFORMUSEWITH', 'MXAPIINSPRESULT', 'MXAPIINTOBJECT', 'MXAPIINVBAL', 'MXAPIINVENTORY', 'MXAPIINVISSUE', 'MXAPIINVOICE', 'MXAPIINVPICKLIST', 'MXAPIINVRES', 'MXAPIINVRESVIEW', 'MXAPIINVUSE', 'MXAPIINVVENDOR', 'MXAPIITEM', 'MXAPIITEMCONDITION', 'MXAPIJOBPLAN', 'MXAPIKPIHISTORY', 'MXAPIKPIMAIN', 'MXAPILABEL', 'MXAPILABOR', 'MXAPILABORCRAFTRATE', 'MXAPILABTRANS', 'MXAPILAUNCHPOINT', 'MXAPILBSLOCATION', 'MXAPILBSLOCHISTORY', 'MXAPILICAPPACCESS', 'MXAPILICENSE', 'MXAPILINEARREFMETHOD', 'MXAPILOCANCESTOR', 'MXAPILOCATION', 'MXAPILOCATIONMETER', 'MXAPILOCATIONS', 'MXAPILOCHIERARCHY', 'MXAPILOCRRPLAN', 'MXAPILOCSYS', 'MXAPILOCSYSTEM', 'MXAPIMAPMANAGER', 'MXAPIMASTERPM', 'MXAPIMATUSETRANSVIEW', 'MXAPIMEASUREMENTCBM', 'MXAPIMEASUREUNIT', 'MXAPIMESSAGE', 'MXAPIMETER', 'MXAPIMETERDATA', 'MXAPIMETERGROUP', 'MXAPIMOBPROPGROUP', 'MXAPIMOBSHAREDDEVICE', 'MXAPIMPFORMULA', 'MXAPIMR', 'MXAPIMSGHUB', 'MXAPIMSGREPROC', 'MXAPIMSGTRACK', 'MXAPIMVI', 'MXAPINOTIFICATION', 'MXAPINUMERICDOMAIN', 'MXAPIOAUTHCLIENT', 'MXAPIOBJECT', 'MXAPIOBJECTAUTH', 'MXAPIOBJFORMULA', 'MXAPIOPERLOC', 'MXAPIOPERLOCINT', 'MXAPIOPERSTORELOC', 'MXAPIORGANIZATION', 'MXAPIOSEVENT', 'MXAPIOSLCQUERY', 'MXAPIPC', 'MXAPIPERGROUP', 'MXAPIPERSON', 'MXAPIPERSONALIZATION', 'MXAPIPERSONGROUP', 'MXAPIPERSONGROUPVIEW', 'MXAPIPERUSER', 'MXAPIPLUSCASSETSTAT', 'MXAPIPM', 'MXAPIPO', 'MXAPIPOFANALYSIS', 'MXAPIPOFUSED', 'MXAPIPR', 'MXAPIPREGENEXECLOG', 'MXAPIPROBLEM', 'MXAPIPROJ', 'MXAPIPROP', 'MXAPIQTOSLCREF', 'MXAPIQUERY', 'MXAPIQUERYRECSTATE', 'MXAPIQUERYTEMPLATE', 'MXAPIRCVROTITM', 'MXAPIRECEIPT', 'MXAPIRESOURCEAUTH', 'MXAPIRETURN', 'MXAPIROUTES', 'MXAPISCORINGRANGE', 'MXAPISERADDRESS', 'MXAPISERVITEM', 'MXAPISITE', 'MXAPISKDODMERUN', 'MXAPISKDPROJACTIVITY', 'MXAPISKDPROJECT', 'MXAPISOLUTION', 'MXAPISR', 'MXAPISRGUESTDOMAIN', 'MXAPISRGUESTEMAIL', 'MXAPISRGUESTPOLICY', 'MXAPISRVAD', 'MXAPISRWATCHLIST', 'MXAPISRWATCHLISTDTS', 'MXAPISTATEMANAGER', 'MXAPISTORELOC', 'MXAPISYNONYMDOMAIN', 'MXAPITABLEDOMAIN', 'MXAPITKCLASS', 'MXAPITKSRVAD', 'MXAPITKTEMPLATE', 'MXAPITOOLINV', 'MXAPITOOLITEM', 'MXAPIVENDOR', 'MXAPIVENDORMSTR', 'MXAPIVIEWMANAGER', 'MXAPIVMUSEROPTIONS', 'MXAPIWEATHERGROUP', 'MXAPIWEATHERORG', 'MXAPIWEATHERPRODUCT', 'MXAPIWEATHERZONE', 'MXAPIWFASSIGNMENT', 'MXAPIWFPROCESS', 'MXAPIWMASSIGNMENT', 'MXAPIWO', 'MXAPIWOASSETTYPE', 'MXAPIWOCHANGESTATUS', 'MXAPIWODETAIL', 'MXAPIWOHIER', 'MXAPIWOINT', 'MXAPIWORKLOG', 'MXAPIWORKQUEUE', 'MXAPIWORKSCAPELAYOUT', 'MXAPIWORKTYPE', 'MXAPIWORKVIEW', 'MXAPIWORKZONE', 'MXAPIWOSDATASET', 'MXAPIWOSRVAD', 'MXAPIWPEDITSETTING'])

MAX_PAGE_SIZE = 100


API_KEY = os.getenv("MAXIMO_API_KEY", "").strip()

HEADERS = {
    "apikey": API_KEY,
    "Accept": "application/json",
}


# ---------------------------------------------------------------------------
# Default fields returned for common Maximo objects
# ---------------------------------------------------------------------------

DEFAULT_SELECTS = {
    "MXAPIWO": (
        "wonum,description,status,"
        "assetnum,location,wopriority,worktype"
    ),
    "MXAPIASSET": (
        "assetnum,description,status,"
        "location,siteid,parent"
    ),
    "MXAPITOOLINV": (
        "itemnum,description,status,"
        "location,siteid,itemtype"
    ),
    "MXAPIITEM": (
        "itemnum,description,status,"
        "itemtype,commodity"
    ),
    "MXAPICONTRACT": (
        "contractnum,description,status,"
        "vendor,startdate,enddate"
    ),
    "MXAPIPO": (
        "ponum,description,status,"
        "vendor,orderdate,requireddate"
    ),
    "MXAPILOCATIONS": (
        "location,description,status,"
        "siteid,parent"
    ),
    "MXAPIPERSON": (
        "personid,displayname,status,"
        "title,primaryphone,primaryemail"
    ),
    "MXAPIPM": (
        "pmnum,description,status,"
        "assetnum,location,siteid,frequency"
    ),
    "MXAPIMASTERPM": (
        "pmnum,description,status,"
        "frequency,frequencyunit"
    ),
}


# ---------------------------------------------------------------------------
# Exceptions
# ---------------------------------------------------------------------------

class MaximoConfigurationError(RuntimeError):
    """Raised when the Maximo API key is not configured."""


# ---------------------------------------------------------------------------
# Internal validation helpers
# ---------------------------------------------------------------------------

def _validate_select(select):
    """Validate an OSLC select field list."""

    if not isinstance(select, str) or not select.strip():
        raise ValueError("Maximo select must be a non-empty field list")

    selected_fields = [
        field.strip()
        for field in select.split(",")
    ]

    if select.strip() == "*":
        return

    if any(
        not re.fullmatch(
            r"[A-Za-z][A-Za-z0-9_]*",
            field
        )
        for field in selected_fields
    ):
        raise ValueError("Maximo select contains an invalid field")


def _validate_identifier_field(identifier_field):
    """Validate a Maximo identifier field name."""

    if not isinstance(identifier_field, str):
        raise ValueError("Maximo identifier field must be a string")

    if not re.fullmatch(
        r"[A-Za-z][A-Za-z0-9_]*",
        identifier_field
    ):
        raise ValueError("Invalid Maximo identifier field")


def _validate_identifier(identifier):
    """Validate a Maximo identifier value."""

    if not isinstance(identifier, str) or not identifier:
        raise ValueError("Maximo identifier must be a non-empty string")

    if not re.fullmatch(
        r"[A-Za-z0-9._-]+",
        identifier
    ):
        raise ValueError("Invalid Maximo identifier")


# ---------------------------------------------------------------------------
# Generic Maximo object request
# ---------------------------------------------------------------------------

def get_object(
    object_name,
    page_size=5,
    where=None,
    page_number=None,
    select=None,
    order_by=None,
):
    """
    Fetch records from a Maximo OSLC object.

    Args:
        object_name: Maximo API object name, e.g. MXAPIWO.
        page_size: Number of records to request.
        where: Optional Maximo OSLC where clause.
        page_number: Optional page number.
        select: Optional comma-separated list of fields.
        order_by: Optional Maximo OSLC order-by expression.

    Returns:
        Parsed JSON response from Maximo.
    """

    # Validate object name.
    if object_name not in MAXIMO_OBJECTS:
        raise ValueError(
            "Maximo object name is not in apis.txt"
        )

    # Validate page size.
    if (
        not isinstance(page_size, int)
        or not 1 <= page_size <= MAX_PAGE_SIZE
    ):
        raise ValueError(
            f"Maximo page size must be between "
            f"1 and {MAX_PAGE_SIZE}"
        )

    # Validate page number.
    if page_number is not None and (
        not isinstance(page_number, int)
        or page_number < 1
    ):
        raise ValueError(
            "Maximo page number must be a positive integer"
        )

    # Validate select.
    if select is not None:
        _validate_select(select)

    # Validate order-by.
    if order_by is not None:
        if not isinstance(order_by, str) or not order_by.strip():
            raise ValueError("Maximo order_by must be a non-empty field name")
        order_fields = [part.strip() for part in order_by.split(",") if part.strip()]
        if not order_fields or any(
            not re.fullmatch(
                r"[A-Za-z][A-Za-z0-9_]*(?:\s+(?:asc|desc))?",
                field,
                re.IGNORECASE
            )
            for field in order_fields
        ):
            raise ValueError("Maximo order_by contains an invalid field")

    # Validate API key.
    if not API_KEY:
        raise MaximoConfigurationError(
            "Set MAXIMO_API_KEY in the project .env file"
        )

    # Build URL.
    url = f"{BASE_URL}/api/os/{object_name}"

    # Base parameters.
    params = {
        "lean": "1",
        "oslc.pageSize": page_size,
    }

    # -----------------------------------------------------------------------
    # Select fields
    # -----------------------------------------------------------------------

    if select is not None:
        params["oslc.select"] = select.strip()

    elif object_name in DEFAULT_SELECTS:
        params["oslc.select"] = DEFAULT_SELECTS[object_name]

    else:
        params["oslc.select"] = "*"

    # -----------------------------------------------------------------------
    # Optional where clause
    # -----------------------------------------------------------------------

    if where:
        params["oslc.where"] = where

    # -----------------------------------------------------------------------
    # Optional pagination
    # -----------------------------------------------------------------------

    if page_number is not None:
        params["pageno"] = page_number

    # -----------------------------------------------------------------------
    # Optional sorting
    # -----------------------------------------------------------------------

    if order_by is not None:
        params["oslc.orderBy"] = order_by.strip()

    # -----------------------------------------------------------------------
    # Execute request
    # -----------------------------------------------------------------------

    started_at = time.perf_counter()

    print(
        f"[Maximo API] GET {url} params={params}"
    )

    try:
        response = requests.get(
            url,
            headers=HEADERS,
            params=params,
            verify=False,
            timeout=30,
        )

        # Maximo may return HTTP 400 when an OSLC select contains an
        # attribute that is not exposed by a particular object structure.
        # Do not interpret that as an authorization failure. Retry once
        # without oslc.select so the object can still be queried with the
        # same credentials.
        if response.status_code == 400 and "oslc.select" in params:
            print(
                "[Maximo API] HTTP 400 with oslc.select; "
                "retrying without explicit select"
            )
            fallback_params = dict(params)
            fallback_params.pop("oslc.select", None)
            response = requests.get(
                url,
                headers=HEADERS,
                params=fallback_params,
                verify=False,
                timeout=30,
            )

    except requests.RequestException as error:
        elapsed = time.perf_counter() - started_at

        print(
            f"[Maximo API] Request failed: "
            f"{type(error).__name__} "
            f"after {elapsed:.2f}s"
        )

        raise

    elapsed = time.perf_counter() - started_at

    print(
        f"[Maximo API] HTTP {response.status_code} "
        f"in {elapsed:.2f}s"
    )

    if response.status_code >= 400:
        print("MAXIMO ERROR BODY:")
        print(response.text)

    # Raise HTTPError for 4xx/5xx responses.
    response.raise_for_status()

    # Parse JSON.
    result = response.json()

    # Log record count.
    members = result.get("member", [])

    print(
        f"[Maximo API] Records returned: "
        f"{len(members)}"
    )

    return result

def change_tool_status(
    itemnum: str | None = None,
    status: str = "NOTACTIVE"
):
    payload = {
        "status": status
    }

    if itemnum:
        return update_tool(itemnum, payload)

    raise ValueError(
        "Bulk update implementation required"
    )
# ---------------------------------------------------------------------------
# Count records
# ---------------------------------------------------------------------------

def get_object_count(object_name, where=None):
    """
    Count every matching Maximo record by following pagination metadata.

    Args:
        object_name: Maximo API object name.
        where: Optional Maximo OSLC where clause.

    Returns:
        Total number of matching records.
    """

    total = 0
    page_number = 1

    while True:
        result = get_object(
            object_name=object_name,
            page_size=MAX_PAGE_SIZE,
            where=where,
            page_number=page_number,
        )
          
        members = result.get("member", [])

        total += len(members)

        response_info = (
            result.get("responseInfo")
            or result.get("responseinfo")
            or {}
        )

        next_page = (
            response_info.get("nextPage")
            or response_info.get("nextpage")
        )

        # No next page means we are finished.
        if not next_page:
            return total

        # Prevent an infinite loop if Maximo returns
        # a next-page link with no records.
        if not members:
            raise RuntimeError(
                "Maximo returned an empty page "
                "with a next-page link"
            )

        # Maximo may return either:
        #   {"href": "..."}
        # or directly:
        #   "https://..."
        next_href = (
            next_page.get("href")
            if isinstance(next_page, dict)
            else next_page
        )

        if not isinstance(next_href, str):
            raise RuntimeError(
                "Maximo returned an invalid next-page link"
            )

        # Extract pageno from the next-page URL.
        query = parse_qs(
            urlsplit(next_href).query
        )

        next_page_number = query.get(
            "pageno",
            [page_number + 1]
        )[0]

        try:
            next_page_number = int(next_page_number)

        except (TypeError, ValueError) as error:
            raise RuntimeError(
                "Maximo returned an invalid "
                "next-page number"
            ) from error

        if next_page_number <= page_number:
            raise RuntimeError(
                "Maximo returned a non-advancing "
                "next-page link"
            )

        page_number = next_page_number


# ---------------------------------------------------------------------------
# Fetch one record
# ---------------------------------------------------------------------------

def get_record(
    object_name,
    identifier_field,
    identifier,
    select=None,
):
    """
    Fetch one Maximo record by identifier field and value.

    Args:
        object_name: Maximo API object name.
        identifier_field: Field used for lookup.
        identifier: Identifier value.
        select: Optional field list.

    Returns:
        Maximo API JSON response.
    """

    _validate_identifier_field(identifier_field)
    _validate_identifier(identifier)

    request_options = {}

    if select is not None:
        request_options["select"] = select

    return get_object(
        object_name=object_name,
        page_size=1,
        where=f'{identifier_field}="{identifier}"',
        **request_options,
    )


# ---------------------------------------------------------------------------
# Work Order
# ---------------------------------------------------------------------------

def get_work_order(wonum, select=None):
    """Fetch one work order by its Maximo work-order number."""

    return get_record(
        "MXAPIWO",
        "wonum",
        wonum,
        select=select,
    )


# ---------------------------------------------------------------------------
# Asset
# ---------------------------------------------------------------------------

def get_asset(assetnum, select=None):
    """Fetch one asset by its Maximo asset number."""

    return get_record(
        "MXAPIASSET",
        "assetnum",
        assetnum,
        select=select,
    )


# ---------------------------------------------------------------------------
# Contract
# ---------------------------------------------------------------------------

def get_contract(contractnum, select=None):
    """Fetch one contract by its contract number."""

    return get_record(
        "MXAPICONTRACT",
        "contractnum",
        contractnum,
        select=select
    )


# ---------------------------------------------------------------------------
# Tool Inventory
# ---------------------------------------------------------------------------

def get_tool(itemnum, select=None):
    """Fetch one stocked tool by its Maximo item number."""

    return get_record(
        "MXAPITOOLINV",
        "itemnum",
        itemnum,
        select=select,
    )


# ---------------------------------------------------------------------------
# Inventory Item
# ---------------------------------------------------------------------------

def get_item(itemnum, select=None):
    """Fetch one inventory item by its Maximo item number."""

    return get_record(
        "MXAPIITEM",
        "itemnum",
        itemnum,
        select=select,
    )


# ---------------------------------------------------------------------------
# Purchase Order
# ---------------------------------------------------------------------------

def get_purchase_order(ponum, select=None):
    """Fetch one purchase order by its PO number."""

    return get_record(
        "MXAPIPO",
        "ponum",
        ponum,
        select=select,
    )


# ---------------------------------------------------------------------------
# Location
# ---------------------------------------------------------------------------

def get_location(location, select=None):
    """Fetch one location by its Maximo location identifier."""

    return get_record(
        "MXAPILOCATIONS",
        "location",
        location,
        select=select,
    )


# ---------------------------------------------------------------------------
# Person
# ---------------------------------------------------------------------------

def get_person(personid, select=None):
    """Fetch one person by their Maximo person identifier."""

    return get_record(
        "MXAPIPERSON",
        "personid",
        personid,
        select=select,
    )

def get_pm(pmnum, select=None):
    """Fetch one PM by PM number."""

    return get_record(
        "MXAPIPM",
        "pmnum",
        pmnum,
        select=select,
    )

def get_master_pm(pmnum, select=None):
    """Fetch one Master PM by PM number."""

    return get_record(
        "MXAPIMASTERPM",
        "pmnum",
        pmnum,
        select=select,
    )

def create_work_order(
    description: str,
    siteid: str,
    location: str | None = None,
    assetnum: str | None = None,
    worktype: str = "CM"
):
    url = f"{BASE_URL}/api/os/MXAPIWO"

    payload = {
        "description": description,
        "siteid": siteid,
        "worktype": worktype
    }

    if location:
        payload["location"] = location

    if assetnum:
        payload["assetnum"] = assetnum

    response = requests.post(
        url,
        headers={
            **HEADERS,
            "Content-Type": "application/json"
        },
        json=payload,
        verify=False,
        timeout=30
    )

    response.raise_for_status()

    return response.json()

def update_work_order(
    wonum: str,
    fields: dict
):
    url = (
        f"{BASE_URL}/api/os/MXAPIWO"
        f'?oslc.where=wonum="{wonum}"'
    )

    response = requests.patch(
        url,
        headers={
            **HEADERS,
            "Content-Type": "application/json",
            "x-method-override": "PATCH"
        },
        json=fields,
        verify=False,
        timeout=30
    )

    response.raise_for_status()

    return {
        "success": True,
        "wonum": wonum
    }

def add_work_log(
    wonum: str,
    log_text: str
):
    url = (
        f"{BASE_URL}/api/os/MXAPIWO"
        f'?oslc.where=wonum="{wonum}"'
    )

    payload = {
        "worklog": [
            {
                "description": "Chatbot Update",
                "description_longdescription": log_text
            }
        ]
    }

    response = requests.patch(
        url,
        headers={
            **HEADERS,
            "Content-Type": "application/json"
        },
        json=payload,
        verify=False,
        timeout=30
    )

    response.raise_for_status()

    return {
        "success": True
    }

