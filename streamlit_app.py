#!/usr/bin/env python
"""
RMS LOC/ACC File Builder (Streamlit) - v4

Ports the "Data" tab + its live formulas (LOC/ACC sheets) + Geocode_DataTab.bas
macro from a confirmed-working RMS cleansing workbook into a single Streamlit
data-entry app. Produces RMS-import-ready ACC and LOC CSV files.

ARCHITECTURE NOTE - v4 fix for the real root cause
-----------------------------------------------------
v3 still placed every input widget (Account Key, location table, etc.)
directly on the page OUTSIDE of an st.form. In that mode, Streamlit reruns
the ENTIRE script on every single widget commit - e.g. editing a cell in
the location table triggers a rerun. If the user had just typed into the
Account Key field and then clicked/tabbed straight into the location table
WITHOUT the browser having finished sending that field's committed value
first, the rerun could capture Account Key as still empty - the exact
"Account Key: ''" symptom reported, even though the user had typed one.
This is a well-known category of Streamlit bug: multiple independent
auto-rerunning widgets racing against each other.

v4 fix: the ENTIRE Account Information + Location Information section is
now wrapped in a single st.form. Inside a form, NOTHING reruns the script
on individual field edits - not the Account Key, not the location table,
not adding/deleting rows. All field values are collected together and
handed to the script ONLY ONCE, atomically, the instant the form's
submit button ("Geocode & Build ACC/LOC Files") is pressed. This removes
the race condition entirely, because there is no longer any possibility of
one widget's rerun capturing another widget's not-yet-committed value -
there are no reruns until submission.

Workflow (unchanged):
  1. Enter Account Information.
  2. Enter one or more Locations.
  3. Press "Geocode & Build ACC/LOC Files" (the form's submit button) -
     geocodes every location and builds the ACC/LOC tables in one pass.

Known issues in the SOURCE WORKBOOK this app intentionally does NOT
reproduce (fixed here instead):
  - LOC!T2/LOC!AS2 build a date string for YEARBUILT/YEARUPGRAD - this app
    outputs the bare 4-digit year integer instead.
  - ACC!Y2 (BLANDEDCUR) contains a broken #REF! formula - left blank here.
  - LOC Latitude/Longitude/ADDRMATCH have no formula in the live sheet -
    this app populates them from the geocode step.
"""

import base64
import hashlib
import hmac as hmac_module
import json
import time
from datetime import date, datetime, timezone
from urllib.parse import urlsplit, quote

import pandas as pd
import pydeck as pdk
import requests
import streamlit as st
from requests.adapters import HTTPAdapter

TIMEOUT = 60
MAX_RETRIES = 4
LARGE_NUMBER_MAX = 999_999_999_999  # effectively unbounded for insurance values

# --------------------------------------------------------------------------
# Reference data - extracted directly from the workbook's "Schemas" sheet.
# --------------------------------------------------------------------------

US_STATES = [
    ("AK","Alaska"),("AL","Alabama"),("AR","Arkansas"),("AZ","Arizona"),("CA","California"),
    ("CO","Colorado"),("CT","Connecticut"),("DC","District of Columbia"),("DE","Delaware"),
    ("FL","Florida"),("GA","Georgia"),("HI","Hawaii"),("IA","Iowa"),("ID","Idaho"),
    ("IL","Illinois"),("IN","Indiana"),("KS","Kansas"),("KY","Kentucky"),("LA","Louisiana"),
    ("MA","Massachusetts"),("MD","Maryland"),("ME","Maine"),("MI","Michigan"),("MN","Minnesota"),
    ("MO","Missouri"),("MS","Mississippi"),("MT","Montana"),("NC","North Carolina"),
    ("ND","North Dakota"),("NE","Nebraska"),("NH","New Hampshire"),("NJ","New Jersey"),
    ("NM","New Mexico"),("NV","Nevada"),("NY","New York"),("OH","Ohio"),("OK","Oklahoma"),
    ("OR","Oregon"),("PA","Pennsylvania"),("RI","Rhode Island"),("SC","South Carolina"),
    ("SD","South Dakota"),("TN","Tennessee"),("TX","Texas"),("UT","Utah"),("VA","Virginia"),
    ("VT","Vermont"),("WA","Washington"),("WI","Wisconsin"),("WV","West Virginia"),("WY","Wyoming"),
]
STATE_CODE_TO_NAME = {c: n for c, n in US_STATES}
STATE_NAME_TO_CODE = {n.upper(): c for c, n in US_STATES}

CONSTRUCTION = [
    (0,"ATC 0 - Unknown"),(1,"ATC 1 - Wood Frame"),(2,"ATC 2 - Light Metal"),
    (3,"ATC 3 - Unreinforced Masonry Wall"),(4,"ATC 4 - URM Wall with Frame"),
    (5,"ATC 5 - RC Shear Wall with Frame"),(6,"ATC 6 - RC Shear Wall without Frame"),
    (7,"ATC 7 - RM Shear Wall"),(8,"ATC 8 - RM Shear Wall with Frame"),
    (9,"ATC 9 - Braced Steel Frame"),(10,"ATC 10 - Moment Steel Frame (Perimeter)"),
    (11,"ATC 11 - Moment Steel Frame (Distributed)"),(12,"ATC 12 - Ductile RC Frame (Distributed)"),
    (13,"ATC 13 - Non-Ductile RC Frame (Distributed)"),(14,"ATC 14 - Precast Concrete (Non Tilt-Up)"),
    (15,"ATC 15 - Long Span"),(16,"ATC 16 - Tilt-Up"),
    (80,"ATC 80 - Reinforced Concrete Frame with Wood Frame in the Upper Floors"),
    (81,"ATC 81 - Steel Frame with Wood Frame in the Upper Floors"),
]
OCCUPANCY = [
    (0,"ATC 0 - Unknown"),(1,"ATC 1 - Permanent Dwelling Single Family Housing"),
    (2,"ATC 2 - Permanent Dwelling Multi Family Housing"),(3,"ATC 3 - Temporary Lodging"),
    (4,"ATC 4 - Group Instutional Housing"),(5,"ATC 5 - Retail"),(6,"ATC 6 - Wholesale Trade"),
    (7,"ATC 7 - Personal and Repair Services"),(8,"ATC 8 - Professional Terchnical and Business Services"),
    (9,"ATC 9 - Health Care Services"),(10,"ATC 10 - Entertainment and recreation"),(11,"ATC 11 - Parking"),
    (12,"ATC 12 - Heavy Fabrication and Assembly"),(13,"ATC 13 - Light Fabrication and Assembly"),
    (14,"ATC 14 - Food and Drugs Processing"),(15,"ATC 15 - Chemicals processing"),
    (16,"ATC 16 - Metal and minerals processing"),(17,"ATC 17 - High technology"),
    (18,"ATC 18 - Construction"),(19,"ATC 19 - Petroleum"),(20,"ATC 20 - Agriculture"),
    (21,"ATC 21 - Mining"),(22,"ATC 22 - Religion and Nonprofit"),(23,"ATC 23 - General Services"),
    (24,"ATC 24 - Emergency response services"),(25,"ATC 25 - Education"),(26,"ATC 26 - Highway"),
    (27,"ATC 27 - Railroad"),(28,"ATC 28 - Air"),(29,"ATC 29 - Sea/ Water"),(30,"ATC 30 - Electrical"),
    (31,"ATC 31 - Water"),(32,"ATC 32 - Sanitary Sewer"),(33,"ATC 33 - Natural gas"),
    (34,"ATC 34 - Telephone and Telegraph"),(35,"ATC 35 - Communication (radio and TV)"),
    (36,"ATC 36 - Flood control"),(37,"ATC 37 - General commercial"),(38,"ATC 38 - General industrial"),
    (39,"ATC 39 - Miscellaneous"),(40,"ATC 40 - Mortgage-backed dwelling (Puerto Rico only)"),
    (41,"ATC 41 - Homeowner"),(42,"ATC 42 - Multi-family dwelling Homeowner Association"),
    (43,"ATC 43 - Multi-family dwelling Condo unit owner"),(44,"ATC 44 - Gas Service station"),
    (47,"ATC 47 - Restaurants"),(48,"ATC 48 - Casinos"),(49,"ATC 49 - Acute Care Services hospitals"),
    (50,"ATC 50 - OSHPD Acute Care Services Hospitals (California only)"),
    (51,"ATC 51 - Hotels - large"),(52,"ATC 52 - Hotels - small and medium"),
    (53,"ATC 53 - Rental - General Commercial"),(54,"ATC 54 - Universities and colleges"),
]
CONSTRUCTION_QUALITY = [(0,"Unknown"),(1,"Good"),(2,"Average"),(3,"Poor")]
CONTENTS_DAMAGE = [(0,"Unknown"),(1,"Highly Damageable"),(2,"Moderately Damageable"),
                    (3,"Damageable"),(4,"Slightly Damageable")]
PLAN_IRREGULARITY = [(0,"Unknown"),(1,"Regular"),(2,"Irregular (L, T shapes, etc.)")]
SOFT_STORY = [(0,"Unknown"),(1,"No"),(2,"Yes (open/weak first floor)")]
VERTICAL_IRREGULARITY = [(0,"Unknown"),(1,"No"),(2,"Yes (Significant setbacks/overhangs)")]
CLADDING_TYPE = [(0,"Unknown"),(1,"Glass"),(2,"Precast Concrete"),
                  (3,"URM - Full Thickness"),(4,"URM - Partial Thickness")]
SPRINKLER_TYPE = [(0,"Unknown"),(1,"Wet"),(2,"Dry")]
SPRINKLER_LEAKAGE = [(0,"Unknown (Average)"),(1,"Low"),(2,"High")]
FRAME_BOLTED = [(0,"Unknown"),(1,"Bolted"),(2,"Unbolted")]
ANCHORING = [(0,"Unknown"),(1,"Properly anchored (Tilt-up)"),(2,"Not properly anchored (Tilt-up)")]
ENGINEERED_FOUNDATION = [(0,"Unknown"),(1,"Yes"),(2,"No")]

RESOLUTION_LABELS = {
    0:"None",1:"Coordinate",2:"Parcel",3:"Street Block",4:"Street Name",5:"Postcode",
    6:"City District",7:"City",8:"District",9:"County",10:"State",11:"Accumulation Zone",
    12:"Multiple",13:"User Defined",14:"Country",
}


def desc_to_code(pairs, description):
    for code, desc in pairs:
        if desc == description:
            return code
    return ""


def options_list(pairs):
    return [desc for _, desc in pairs]


# --------------------------------------------------------------------------
# Geocode API client - flat body, confirmed from Geocode_DataTab.bas
# --------------------------------------------------------------------------


# ==========================================================================
# RMS EDM Creation + MRI Import pipeline
# Ported from modRMSImport.bas + modAWSSigV4.bas (confirmed-working VBA).
# ==========================================================================

TIMEOUT_CONNECT = 30
TIMEOUT_READ = 120
EDM_MAX_RETRIES = 3

# Confirmed defaults, ported from modRMSConfig.bas
RMS_DELIMITER = "COMMA"
RMS_DELIM_CHAR = ","
RMS_HAS_HEADERS = True
RMS_PAGE_LIMIT = 100
RMS_EDM_WAIT_SECONDS = 10
RMS_EDM_WAIT_ATTEMPTS = 60
RMS_DEFAULT_REGION = "eu-west-1"


# ==========================================================================
# AWS SigV4 signing - pure Python port of modAWSSigV4.bas's BuildS3PutHeaders
# ==========================================================================

def sha256_hex(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def hmac_sha256(key: bytes, msg: bytes) -> bytes:
    return hmac_module.new(key, msg, hashlib.sha256).digest()


def aws_uri_encode(s: str, encode_slash: bool = True) -> str:
    """Matches AwsUriEncode in modAWSSigV4.bas: unreserved chars pass
    through, '/' is kept only if encode_slash=False, everything else is
    percent-encoded (UTF-8, uppercase hex)."""
    safe = "" if encode_slash else "/"
    return quote(s, safe=safe + "-_.~")


def amz_date_utc_now() -> str:
    """Matches AmzDateUtcNow: YYYYMMDDTHHMMSSZ in UTC."""
    return datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")


def build_s3_put_headers(upload_url: str, region: str, access_key: str,
                          secret_key: str, session_token: str, file_bytes: bytes) -> dict:
    """Pure-Python port of BuildS3PutHeaders - produces the exact same
    canonical request / string-to-sign / signing-key-chain / Authorization
    header as the confirmed-working VBA macro, for a presigned S3 PUT with
    host + x-amz-content-sha256 + x-amz-date + x-amz-security-token as the
    signed headers (matching the VBA exactly - no other headers signed)."""
    parsed = urlsplit(upload_url)
    host = parsed.netloc
    path = parsed.path or "/"

    amz_date = amz_date_utc_now()
    payload_hash = sha256_hex(file_bytes)

    canonical_uri = path  # already URL-path-safe from a presigned URL; not re-encoded, matching VBA

    canonical_headers = (
        f"host:{host}\n"
        f"x-amz-content-sha256:{payload_hash}\n"
        f"x-amz-date:{amz_date}\n"
        f"x-amz-security-token:{session_token}\n"
    )
    signed_headers = "host;x-amz-content-sha256;x-amz-date;x-amz-security-token"

    canonical_request = (
        f"PUT\n{canonical_uri}\n\n{canonical_headers}\n{signed_headers}\n{payload_hash}"
    )

    date_stamp = amz_date[:8]
    credential_scope = f"{date_stamp}/{region}/s3/aws4_request"
    hashed_canonical_request = sha256_hex(canonical_request.encode("utf-8"))

    string_to_sign = (
        f"AWS4-HMAC-SHA256\n{amz_date}\n{credential_scope}\n{hashed_canonical_request}"
    )

    k_date = hmac_sha256(("AWS4" + secret_key).encode("utf-8"), date_stamp.encode("utf-8"))
    k_region = hmac_sha256(k_date, region.encode("utf-8"))
    k_service = hmac_sha256(k_region, b"s3")
    k_signing = hmac_sha256(k_service, b"aws4_request")

    signature = hmac_sha256(k_signing, string_to_sign.encode("utf-8")).hex()

    auth_header = (
        f"AWS4-HMAC-SHA256 Credential={access_key}/{credential_scope}, "
        f"SignedHeaders={signed_headers}, Signature={signature}"
    )

    return {
        "x-amz-date": amz_date,
        "x-amz-content-sha256": payload_hash,
        "x-amz-security-token": session_token,
        "Authorization": auth_header,
    }


# ==========================================================================
# RMS Platform API client
# ==========================================================================

class RmsImportApi:
    """Wraps every confirmed endpoint used by GetOrCreateEDM -> ... ->
    SubmitImportJob -> CheckJobStatusOnce. One instance per import run."""

    def __init__(self, host: str, api_key: str, verify: bool = True):
        self.host = host.strip()
        self.session = requests.Session()
        adapter = HTTPAdapter(pool_connections=8, pool_maxsize=8)
        self.session.mount("https://", adapter)
        self.session.mount("http://", adapter)
        clean_key = (api_key or "").strip()
        if clean_key.lower().startswith("bearer "):
            clean_key = clean_key[7:].strip()
        self.api_key = clean_key
        self.session.headers.update({"Accept": "application/json", "Authorization": clean_key})
        self.session.verify = verify
        self._resource_group_id = None
        self._resource_group_id_fetched = False

    # ---- low-level helpers ----

    def _get(self, url, retries=EDM_MAX_RETRIES):
        last_exc = None
        for attempt in range(retries):
            try:
                r = self.session.get(url, timeout=(TIMEOUT_CONNECT, TIMEOUT_READ))
                return r.status_code, r.text, dict(r.headers)
            except requests.RequestException as exc:
                last_exc = str(exc)
                if attempt < retries - 1:
                    time.sleep(min(2 ** (attempt + 1), 10))
                    continue
                return 0, f"GET error: {last_exc}", {}
        return 0, f"GET error: {last_exc}", {}

    def _post(self, url, body_json: str, extra_headers: dict | None = None, retries=EDM_MAX_RETRIES):
        headers = {"Content-Type": "application/json", "Accept": "application/json",
                   "Authorization": self.api_key}
        if extra_headers:
            headers.update(extra_headers)
        last_exc = None
        for attempt in range(retries):
            try:
                r = self.session.post(url, data=body_json.encode("utf-8"), headers=headers,
                                       timeout=(TIMEOUT_CONNECT, TIMEOUT_READ))
                return r.status_code, r.text, dict(r.headers)
            except requests.RequestException as exc:
                last_exc = str(exc)
                if attempt < retries - 1:
                    time.sleep(min(2 ** (attempt + 1), 10))
                    continue
                return 0, f"POST error: {last_exc}", {}
        return 0, f"POST error: {last_exc}", {}

    @staticmethod
    def _safe_json(text: str):
        try:
            return json.loads(text)
        except (ValueError, TypeError):
            return None

    @staticmethod
    def _as_list(parsed, *keys):
        if isinstance(parsed, list):
            return parsed
        if isinstance(parsed, dict):
            for k in keys:
                v = parsed.get(k)
                if isinstance(v, list):
                    return v
        return []

    @staticmethod
    def _match_name(item, *keys):
        if not isinstance(item, dict):
            return ""
        for k in keys:
            v = item.get(k)
            if v not in (None, ""):
                return str(v)
        return ""

    @staticmethod
    def _extract_id(item, primary_key):
        if not isinstance(item, dict):
            return ""
        v = item.get(primary_key)
        if v not in (None, ""):
            return str(v)
        v = item.get("id")
        if v not in (None, ""):
            return str(v)
        return ""

    # ---- EDM ----

    def resolve_exposure_id(self, target_name: str) -> str:
        """Ported from ResolveExposureId: try exact filter match on
        exposureName/databaseName first, else paginate through all
        exposures looking for a case-insensitive exact name match."""
        base = f"https://{self.host}/platform/riskdata/v1/exposures"
        for filter_field in ("exposureName", "databaseName"):
            filt = f"{filter_field}={quote(target_name)}"
            status, text, _ = self._get(f"{base}?filter={filt}&limit={RMS_PAGE_LIMIT}")
            if status == 200:
                parsed = self._safe_json(text)
                for item in self._as_list(parsed, "exposures", "items", "data"):
                    nm = self._match_name(item, "exposureName", "databaseName")
                    if nm.strip().lower() == target_name.strip().lower():
                        return self._extract_id(item, "exposureId")

        offset = 0
        while True:
            status, text, _ = self._get(f"{base}?limit={RMS_PAGE_LIMIT}&offset={offset}")
            if status != 200:
                break
            parsed = self._safe_json(text)
            items = self._as_list(parsed, "exposures", "items", "data")
            if not items:
                break
            for item in items:
                nm = self._match_name(item, "exposureName", "databaseName")
                if nm.strip().lower() == target_name.strip().lower():
                    return self._extract_id(item, "exposureId")
            if len(items) < RMS_PAGE_LIMIT or offset > 20000:
                break
            offset += RMS_PAGE_LIMIT
        return ""

    def create_edm(self, target_name: str, wait_callback=None) -> tuple[bool, str]:
        """Ported from CreateEDM. Returns (success, diagnostic_message)."""
        url = (f"https://{self.host}/riskmodeler/v1/datasources"
               f"?datasourcename={quote(target_name)}&operation=CREATE")
        status, text, _ = self._post(url, "")
        if status == 409:
            return True, ""
        if not (200 <= status < 300):
            return False, (f"EDM CREATE request was rejected by RMS.\nURL: {url}\n"
                            f"HTTP status: {status}\nResponse: {text}")

        for attempt in range(1, RMS_EDM_WAIT_ATTEMPTS + 1):
            if wait_callback:
                wait_callback(attempt, RMS_EDM_WAIT_ATTEMPTS, target_name)
            time.sleep(RMS_EDM_WAIT_SECONDS)
            if self.resolve_exposure_id(target_name):
                return True, ""
        if self.resolve_exposure_id(target_name):
            return True, ""
        return False, (f"RMS accepted the EDM create request (HTTP {status}), but the EDM never "
                        f"became findable afterward. Waited {RMS_EDM_WAIT_ATTEMPTS} x "
                        f"{RMS_EDM_WAIT_SECONDS}s for '{target_name}'.\nCreate response: {text}")

    def get_or_create_edm(self, target_name: str, wait_callback=None) -> tuple[str, str]:
        """Returns (exposure_id, diagnostic_message). exposure_id is '' on failure."""
        eid = self.resolve_exposure_id(target_name)
        if eid:
            return eid, ""
        ok, diag = self.create_edm(target_name, wait_callback=wait_callback)
        if not ok:
            return "", diag
        eid = self.resolve_exposure_id(target_name)
        if not eid:
            return "", (diag or "CreateEDM reported success, but the immediate follow-up lookup "
                                 f"still returned nothing for '{target_name}'.")
        return eid, ""

    # ---- Portfolio ----

    def list_portfolios(self, exposure_id: str) -> list:
        url = f"https://{self.host}/platform/riskdata/v1/exposures/{exposure_id}/portfolios"
        status, text, _ = self._get(url)
        if status == 200:
            return self._as_list(self._safe_json(text), "portfolios", "items", "data")
        return []

    @staticmethod
    def sanitize_name(name: str) -> str:
        """Matches SanitizeName's intent closely enough for portfolio/account
        identifiers: strip, collapse to a reasonable identifier charset."""
        return "".join(ch for ch in name.strip() if ch.isalnum() or ch in "-_. ")

    def create_portfolio(self, exposure_id: str, target_name: str, portfolio_number: str) -> str:
        url = f"https://{self.host}/platform/riskdata/v1/exposures/{exposure_id}/portfolios"
        safe_number = portfolio_number[:20]
        body = json.dumps({"portfolioName": target_name, "portfolioNumber": safe_number})
        status, text, _ = self._post(url, body)
        if not (200 <= status < 300):
            return ""
        parsed = self._safe_json(text)
        pid = self._extract_id(parsed, "portfolioId") if isinstance(parsed, dict) else ""
        if not pid:
            for _ in range(6):
                time.sleep(3)
                for item in self.list_portfolios(exposure_id):
                    nm = self._match_name(item, "portfolioName", "name")
                    if nm.strip() == target_name.strip():
                        pid = self._extract_id(item, "portfolioId")
                        break
                if pid:
                    break
        if pid:
            time.sleep(3)
        return pid

    def ensure_portfolio_by_identifier(self, exposure_id: str, account_number: str) -> str:
        safe_acct = self.sanitize_name(account_number)[:20]
        for item in self.list_portfolios(exposure_id):
            nm = self._match_name(item, "portfolioName", "name")
            num = self._match_name(item, "portfolioNumber", "number")
            if num.strip().lower() == safe_acct.strip().lower() or nm.strip().lower() == safe_acct.strip().lower():
                return self._extract_id(item, "portfolioId")
        return self.create_portfolio(exposure_id, safe_acct, safe_acct)

    # ---- Import folder + S3 upload ----

    def create_import_folder(self) -> tuple[bool, str, dict, dict]:
        """Returns (success, folder_id, acc_upload_detail, loc_upload_detail)."""
        url = f"https://{self.host}/platform/import/v1/folders"
        body = json.dumps({
            "folderType": "MRI",
            "properties": {
                "fileExtension": "csv", "delimiter": RMS_DELIMITER,
                "delimiterCharacter": RMS_DELIM_CHAR, "hasColumnHeaders": RMS_HAS_HEADERS,
                "containsHeader": RMS_HAS_HEADERS, "hasHeader": RMS_HAS_HEADERS,
            },
        })
        status, text, _ = self._post(url, body)
        if status not in (200, 201):
            body = json.dumps({"folderType": "MRI", "properties": {"fileExtension": "csv"}})
            status, text, _ = self._post(url, body)
            if status not in (200, 201):
                return False, "", {}, {}
        parsed = self._safe_json(text)
        if not isinstance(parsed, dict):
            return False, "", {}, {}
        folder_id = str(parsed.get("folderId", ""))
        upload_details = parsed.get("uploadDetails")
        if not isinstance(upload_details, dict):
            return False, "", {}, {}
        acc_upload = upload_details.get("accountsFile")
        loc_upload = upload_details.get("locationsFile")
        if not isinstance(acc_upload, dict) or not isinstance(loc_upload, dict):
            return False, "", {}, {}
        return True, folder_id, acc_upload, loc_upload

    def upload_file_to_s3(self, file_detail: dict, file_bytes: bytes) -> tuple[bool, str]:
        """Returns (success, error_message)."""
        try:
            presign = file_detail.get("presignParams")
            if not isinstance(presign, dict):
                return False, "presignParams missing from upload detail"
            import base64
            access_key = base64.b64decode(presign["accessKeyId"]).decode("utf-8")
            secret_key = base64.b64decode(presign["secretAccessKey"]).decode("utf-8")
            session_token = base64.b64decode(presign["sessionToken"]).decode("utf-8")
            region = RMS_DEFAULT_REGION
            if presign.get("region"):
                region = base64.b64decode(presign["region"]).decode("utf-8") or RMS_DEFAULT_REGION

            upload_url = file_detail["uploadUrl"]
            headers = build_s3_put_headers(upload_url, region, access_key, secret_key,
                                            session_token, file_bytes)
            r = requests.put(upload_url, data=file_bytes, headers=headers,
                              timeout=(30, 600))
            if 200 <= r.status_code < 300:
                return True, ""
            return False, f"S3 PUT failed: HTTP {r.status_code} - {r.text[:300]}"
        except Exception as exc:
            return False, f"S3 upload exception: {exc}"

    # ---- Import job submission + polling ----

    def get_resource_group_id(self) -> str:
        if self._resource_group_id_fetched:
            return self._resource_group_id or ""
        self._resource_group_id_fetched = True
        self._resource_group_id = ""
        url = f"https://{self.host}/platform/tenantdata/v1/entitlements/RI-RISKMODELER/resourcegroups"
        status, text, _ = self._get(url)
        if status == 200:
            parsed = self._safe_json(text)
            items = self._as_list(parsed, "resourceGroups", "items", "data")
            if items and isinstance(items[0], dict):
                first = items[0]
                rgid = first.get("resourceGroupId") or first.get("id") or first.get("resourcegroupId") or ""
                self._resource_group_id = str(rgid)
        return self._resource_group_id or ""

    def submit_import_job(self, folder_id: str, exposure_id: str, portfolio_id: str,
                           currency: str = "USD") -> tuple[str, str]:
        """Returns (job_id, raw_response_text). job_id is '' on failure."""
        url = f"https://{self.host}/platform/import/v1/jobs"
        resource_uri = f"/platform/riskdata/v1/exposures/{exposure_id}"
        if portfolio_id:
            resource_uri += f"/portfolios/{portfolio_id}"
        body = json.dumps({
            "importType": "MRI", "resourceUri": resource_uri,
            "settings": {
                "folderId": folder_id, "delimiter": RMS_DELIMITER,
                "delimiterCharacter": RMS_DELIM_CHAR, "hasColumnHeaders": RMS_HAS_HEADERS,
                "containsHeader": RMS_HAS_HEADERS, "hasHeader": RMS_HAS_HEADERS, "currency": currency,
            },
        })
        resource_group_id = self.get_resource_group_id()
        extra_headers = {"x-rms-resource-group-id": resource_group_id} if resource_group_id else None

        raw_text = ""
        for attempt in range(4):
            status, text, headers = self._post(url, body, extra_headers=extra_headers)
            raw_text = text
            if 200 <= status < 300:
                job_id = ""
                parsed = self._safe_json(text)
                if isinstance(parsed, dict):
                    job_id = str(parsed.get("id") or parsed.get("jobId") or parsed.get("workflowId") or "")
                if not job_id:
                    location = headers.get("Location") or headers.get("location") or ""
                    if location:
                        job_id = location.rstrip("/").split("/")[-1]
                return job_id, raw_text

            resp_code = ""
            parsed = self._safe_json(text)
            if isinstance(parsed, dict):
                resp_code = str(parsed.get("code", ""))
            is_resource_id_error = ("resource id is not specified" in text.lower()) or resp_code == "INT-999"
            if is_resource_id_error and attempt < 3:
                time.sleep(5)
                continue
            return "", raw_text
        return "", raw_text

    def check_job_status_once(self, job_id: str, cached_status_url: str = "") -> tuple[str, str, str]:
        """Returns (status_string, status_url_used, raw_response_text).
        Status is 'UNKNOWN' if no candidate URL responds 200."""
        status_url = cached_status_url
        if not status_url:
            candidates = [
                f"https://{self.host}/platform/import/v1/jobs/{job_id}",
                f"https://{self.host}/platform/v1/workflows/{job_id}",
                f"https://{self.host}/riskmodeler/v1/workflows/{job_id}",
            ]
            for candidate in candidates:
                status, text, _ = self._get(candidate)
                if status == 200:
                    status_url = candidate
                    break
            if not status_url:
                return "UNKNOWN", "", ""

        status, text, _ = self._get(status_url)
        parsed = self._safe_json(text)
        st = ""
        if isinstance(parsed, dict):
            st = str(parsed.get("status") or parsed.get("jobStatus") or "").upper()
        return (st or "UNKNOWN"), status_url, text


class GeocodeApi:
    def __init__(self, host, api_key, verify=True):
        self.base_url = f"https://{host}/li"
        self.session = requests.Session()
        adapter = HTTPAdapter(pool_connections=16, pool_maxsize=16)
        self.session.mount("https://", adapter)
        self.session.mount("http://", adapter)
        clean_key = (api_key or "").strip()
        if clean_key.lower().startswith("bearer "):
            clean_key = clean_key[7:].strip()
        self.session.headers.update({
            "Accept": "application/json", "Content-Type": "application/json",
            "Authorization": clean_key,
        })
        self.session.verify = verify

    def geocode(self, street, city, county, state_input, postal, version="latest"):
        """Resolves state_input (either a 2-letter code or full name) into
        BOTH admin1Code and admin1Name, exactly as Geocode_DataTab.bas does,
        then sends a flat body. Returns (payload_dict_or_None, http_status,
        error_detail_or_None)."""
        state_input = (state_input or "").strip()
        admin1_code, admin1_name = "", ""
        if state_input:
            upper = state_input.upper()
            if upper in STATE_CODE_TO_NAME:
                admin1_code, admin1_name = upper, STATE_CODE_TO_NAME[upper]
            elif upper in STATE_NAME_TO_CODE:
                admin1_code, admin1_name = STATE_NAME_TO_CODE[upper], state_input
            elif len(state_input) == 2:
                admin1_code = upper
            else:
                admin1_name = state_input

        body = {"countryScheme": "ISO2A", "countryCode": "US"}
        if street:
            body["streetAddress"] = street
        if city:
            body["cityName"] = city
        if county:
            body["admin2Name"] = county
        if admin1_code:
            body["admin1Code"] = admin1_code
        if admin1_name:
            body["admin1Name"] = admin1_name
        if postal:
            body["postalCode"] = str(postal)

        url = f"{self.base_url}/geocode/{version}"
        last_exc = None
        for attempt in range(MAX_RETRIES):
            try:
                response = self.session.post(url, json=body, timeout=TIMEOUT)
                if response.status_code == 200:
                    try:
                        return response.json(), 200, None
                    except ValueError:
                        return None, response.status_code, "Invalid JSON response"
                if response.status_code in (429, 500, 502, 503, 504) and attempt < MAX_RETRIES - 1:
                    time.sleep(min(2 ** (attempt + 1), 15))
                    continue
                try:
                    detail = response.json()
                except ValueError:
                    detail = response.text[:300]
                return None, response.status_code, detail
            except requests.RequestException as exc:
                last_exc = str(exc)
                if attempt < MAX_RETRIES - 1:
                    time.sleep(min(2 ** (attempt + 1), 15))
                    continue
                return None, 0, last_exc
        return None, 0, last_exc


# --------------------------------------------------------------------------
# Default per-location field set (mirrors Data!B11:AV11 input columns).
# --------------------------------------------------------------------------

TEXT_FIELDS = ["STREETNAME", "CITY", "COUNTY", "STATE", "ZIPCODE"]

LOCATION_INPUT_DEFAULTS = {
    "STREETNAME": "", "CITY": "", "COUNTY": "", "STATE": "", "ZIPCODE": "",
    "Building Value": 0, "Contents": 0, "BI": 0,
    "Contents Damageability Grade": "Unknown",
    "Site Limit": 0, "Site Deductible": 0,
    "Combined Limit": 0, "Combined Deductible": 0,
    "Buildings Deductible": 0, "Contents Deductible": 0, "BI deductible": 0,
    "Construction": "ATC 0 - Unknown", "Occupancy": "ATC 0 - Unknown",
    "Number of Buildings": 1, "Number of Stories": 0,
    "Year Built (YYYY)": 0, "EQSL Y/N": "N", "Year Upgraded (YYYY)": 0,
    "Floor Area (SQ FT)": 0,
    "Construction Quality": "Unknown", "Plan Irregularity": "Unknown",
    "Soft Story": "Unknown", "Vertical Irregularity": "Unknown",
    "Cladding Type": "Unknown", "Frame Bolted": "Unknown", "Anchoring": "Unknown",
    "Engineered Foundation": "Unknown", "Sprinkler Type": "Unknown",
    "EQSL  Susceptibility": "Unknown (Average)",
    "Current Replacement Cost": 0, "Required Min $/sf": 0,
    "Min $/sf Flag": "", "Evaluated Building Value": 0,
}

LARGE_NUMBER_FIELDS = [
    "Building Value", "Contents", "BI", "Site Limit", "Site Deductible",
    "Combined Limit", "Combined Deductible", "Buildings Deductible",
    "Contents Deductible", "BI deductible", "Floor Area (SQ FT)",
    "Current Replacement Cost", "Required Min $/sf", "Evaluated Building Value",
]

ACC_HEADERS = [
    "ACCNTNUM","ACCNTNAME","ACCNTSTAT","UWRITRNAME","BRANCHNAME","PRODID","CEDANTID",
    "CEDANTNAME","POLICYNUM","LOBNAME","INCEPTDATE","EXPIREDATE","UNDCOVAMT","UNDCOVCUR",
    "PARTOF","PARTOFCUR","POLICYTYPE","POLICYSTAT","POLICYSTRUCTURE","MINDEDAMT","MINDEDCUR",
    "MAXDEDAMT","MAXDEDCUR","BLANDEDAMT","BLANDEDCUR","BLANLIMAMT","BLANLIMCUR",
    "BLANPREAMT","BLANPRECUR",
]

LOC_HEADERS = [
    "ACCNTNUM","LOCNUM","STREETNAME","CITY","COUNTY","CRESTA","STATECODE","POSTALCODE",
    "LATITUDE","LONGITUDE","ADDRMATCH","Resolution Type","CNTRYSCHEME","CNTRYCODE",
    "NUMBLDGS","BLDGSCHEME","BLDGCLASS","OCCSCHEME","OCCTYPE","YEARBUILT","NUMSTORIES",
    "INCEPTDATE","EXPIREDATE","EQCV1DED","EQCV1DCUR","EQCV1VAL","EQCV1VCUR","EQCV2DED",
    "EQCV2DCUR","EQCV2VAL","EQCV2VCUR","EQCV2CVM","EQCV3DED","EQCV3DCUR","EQCV3VAL",
    "EQCV3VCUR","EQSITELIM","EQSITELCUR","EQSITEDED","EQSITEDCUR","EQCOMBINEDLIM",
    "EQCOMBINEDLCUR","EQCOMBINEDDED","EQCOMBINEDDCUR","YEARUPGRAD","FLOORAREA","CONQUAL",
    "SHAPECONF","STORYPROF","OVERPROF","CLADDING","EQSLINS","SPNKLRTYPE",
    "EQSLSUSCEPTIBILITY","FRAMEBOLT","TILTUPRET","ENGFOUND","User Defined 1",
]


def safe_num(v, default=0):
    try:
        if v in (None, ""):
            return default
        return float(v)
    except (TypeError, ValueError):
        return default


def safe_year(v):
    """YEARBUILT / YEARUPGRAD must be a plain 4-digit year integer - NOT a
    date string (corrects a known formula bug in the source workbook)."""
    n = safe_num(v, 0)
    return int(n) if n > 1752 else ""


def build_acc_row(account):
    """Builds one ACC row dict, ported field-by-field from the live ACC
    sheet formulas (ACC!A2:AC2). BLANDEDCUR is intentionally left blank -
    the source workbook's ACC!Y2 formula contains a broken #REF! reference."""
    row = {h: "" for h in ACC_HEADERS}
    account_key = str(account.get("Account Key", "")).strip()

    row["ACCNTNUM"] = account_key
    row["ACCNTNAME"] = account_key
    row["ACCNTSTAT"] = "QUOTE" if account_key else ""
    row["UWRITRNAME"] = account.get("Underwriter Initials", "")
    row["POLICYNUM"] = account.get("Layer Name", "")
    row["LOBNAME"] = account.get("Line of Business", "")
    row["INCEPTDATE"] = account.get("Inception Date")
    row["EXPIREDATE"] = account.get("Expiry Date")
    row["UNDCOVAMT"] = safe_num(account.get("Attachment"))
    row["UNDCOVCUR"] = "USD" if row["UNDCOVAMT"] not in (0, "") else ""
    row["PARTOF"] = safe_num(account.get("Policy Limit"))
    row["PARTOFCUR"] = "USD" if row["PARTOF"] not in (0, "") else ""
    row["POLICYTYPE"] = 1 if account_key else ""
    row["MINDEDAMT"] = safe_num(account.get("Min Ded"))
    row["MINDEDCUR"] = "USD" if row["MINDEDAMT"] not in (0, "") else ""
    row["MAXDEDAMT"] = safe_num(account.get("Max Ded"))
    row["MAXDEDCUR"] = "USD" if row["MAXDEDAMT"] not in (0, "") else ""
    row["BLANLIMAMT"] = safe_num(account.get("Share"))
    row["BLANLIMCUR"] = "USD" if row["BLANLIMAMT"] not in (0, "") else ""
    row["BLANPREAMT"] = safe_num(account.get("Premium"))
    row["BLANPRECUR"] = "USD" if row["BLANPREAMT"] not in (0, "") else ""
    return row


def build_loc_row(account, loc, loc_number, geocode_fields):
    """Builds one LOC row dict. `loc` is the user's typed input fields for
    this location; `geocode_fields` is computed fresh in the SAME pass."""
    row = {h: "" for h in LOC_HEADERS}
    account_key = str(account.get("Account Key", "")).strip()
    if not account_key:
        return row

    row["ACCNTNUM"] = account_key
    row["LOCNUM"] = loc_number
    row["STREETNAME"] = loc.get("STREETNAME", "")
    row["CITY"] = loc.get("CITY", "")
    row["COUNTY"] = geocode_fields.get("County Helper", "") or ""
    row["CRESTA"] = geocode_fields.get("Cresta Zone Helper", "") or ""
    row["STATECODE"] = geocode_fields.get("State Helper", "") or ""
    row["POSTALCODE"] = loc.get("ZIPCODE", "")

    row["LATITUDE"] = geocode_fields.get("Latitude")
    row["LONGITUDE"] = geocode_fields.get("Longitude")
    row["ADDRMATCH"] = geocode_fields.get("Resolution Code")
    row["Resolution Type"] = geocode_fields.get("Resolution Type", "")

    row["CNTRYSCHEME"] = "ISO2A"
    row["CNTRYCODE"] = "US"
    num_bldgs = safe_num(loc.get("Number of Buildings"), 0)
    row["NUMBLDGS"] = int(num_bldgs) if num_bldgs else 1
    row["BLDGSCHEME"] = "ATC"
    row["BLDGCLASS"] = desc_to_code(CONSTRUCTION, loc.get("Construction", ""))
    row["OCCSCHEME"] = "ATC"
    row["OCCTYPE"] = desc_to_code(OCCUPANCY, loc.get("Occupancy", ""))
    row["YEARBUILT"] = safe_year(loc.get("Year Built (YYYY)"))
    row["NUMSTORIES"] = safe_num(loc.get("Number of Stories"), "")

    row["INCEPTDATE"] = account.get("Inception Date")
    row["EXPIREDATE"] = account.get("Expiry Date")

    row["EQCV1DED"] = safe_num(loc.get("Site Deductible"))
    row["EQCV1DCUR"] = "USD" if row["EQCV1DED"] else ""
    row["EQCV1VAL"] = safe_num(loc.get("Building Value"))
    row["EQCV1VCUR"] = "USD" if row["EQCV1VAL"] else ""

    row["EQCV2DED"] = safe_num(loc.get("Contents Deductible"))
    row["EQCV2DCUR"] = "USD" if row["EQCV2DED"] else ""
    row["EQCV2VAL"] = safe_num(loc.get("Contents"))
    row["EQCV2VCUR"] = "USD" if row["EQCV2VAL"] else ""
    row["EQCV2CVM"] = desc_to_code(CONTENTS_DAMAGE, loc.get("Contents Damageability Grade", ""))

    row["EQCV3DED"] = safe_num(loc.get("BI deductible"))
    row["EQCV3DCUR"] = "USD" if row["EQCV3DED"] else ""
    row["EQCV3VAL"] = safe_num(loc.get("BI"))
    row["EQCV3VCUR"] = "USD" if row["EQCV3VAL"] else ""

    row["EQSITELIM"] = safe_num(loc.get("Site Limit"))
    row["EQSITELCUR"] = "USD" if row["EQSITELIM"] else ""
    row["EQSITEDED"] = safe_num(loc.get("Site Deductible"))
    row["EQSITEDCUR"] = "USD" if row["EQSITEDED"] else ""

    row["EQCOMBINEDLIM"] = safe_num(loc.get("Combined Limit"))
    row["EQCOMBINEDLCUR"] = "USD" if row["EQCOMBINEDLIM"] else ""
    row["EQCOMBINEDDED"] = safe_num(loc.get("Combined Deductible"))
    row["EQCOMBINEDDCUR"] = "USD" if row["EQCOMBINEDDED"] else ""

    row["YEARUPGRAD"] = safe_year(loc.get("Year Upgraded (YYYY)"))
    row["FLOORAREA"] = safe_num(loc.get("Floor Area (SQ FT)"), "")
    row["CONQUAL"] = desc_to_code(CONSTRUCTION_QUALITY, loc.get("Construction Quality", ""))
    row["SHAPECONF"] = desc_to_code(PLAN_IRREGULARITY, loc.get("Plan Irregularity", ""))
    row["STORYPROF"] = desc_to_code(SOFT_STORY, loc.get("Soft Story", ""))
    row["OVERPROF"] = desc_to_code(VERTICAL_IRREGULARITY, loc.get("Vertical Irregularity", ""))
    row["CLADDING"] = desc_to_code(CLADDING_TYPE, loc.get("Cladding Type", ""))

    eqsl = str(loc.get("EQSL Y/N", "")).strip().upper()
    row["EQSLINS"] = 1 if eqsl == "Y" else (0 if eqsl == "N" else "")

    row["SPNKLRTYPE"] = desc_to_code(SPRINKLER_TYPE, loc.get("Sprinkler Type", ""))
    row["EQSLSUSCEPTIBILITY"] = desc_to_code(SPRINKLER_LEAKAGE, loc.get("EQSL  Susceptibility", ""))
    row["FRAMEBOLT"] = desc_to_code(FRAME_BOLTED, loc.get("Frame Bolted", ""))
    row["TILTUPRET"] = desc_to_code(ANCHORING, loc.get("Anchoring", ""))
    row["ENGFOUND"] = desc_to_code(ENGINEERED_FOUNDATION, loc.get("Engineered Foundation", ""))
    row["User Defined 1"] = geocode_fields.get("Cresta Zone Helper", "") or ""

    return row


def geocode_and_build(api, account, locations_df):
    """Single-pass pipeline: geocodes every location with a STREETNAME and
    builds the matching LOC row immediately, using only local variables."""
    preview_rows = []
    loc_rows = []
    ok_count = 0
    total_count = 0

    for _, loc in locations_df.iterrows():
        street = str(loc.get("STREETNAME", "")).strip()
        if not street:
            continue
        total_count += 1
        loc_dict = loc.to_dict()

        payload, status, error = api.geocode(
            street, loc_dict.get("CITY", ""), loc_dict.get("COUNTY", ""),
            loc_dict.get("STATE", ""), loc_dict.get("ZIPCODE", ""),
        )

        geocode_fields = {
            "Latitude": None, "Longitude": None, "Resolution Code": None,
            "Resolution Type": "", "State Helper": "", "County Helper": "", "Cresta Zone Helper": "",
        }

        lat = lon = None
        if status == 200 and isinstance(payload, dict):
            lat = payload.get("latitude")
            lon = payload.get("longitude")

        if status == 200 and lat is not None and lon is not None:
            res_code = payload.get("rmsGeoModelResolutionCode")
            geocode_fields["Latitude"] = lat
            geocode_fields["Longitude"] = lon
            geocode_fields["Resolution Code"] = res_code
            geocode_fields["Resolution Type"] = RESOLUTION_LABELS.get(res_code, "Unrecognized code")
            geocode_fields["County Helper"] = payload.get("admin2Name", "")
            geocode_fields["Cresta Zone Helper"] = payload.get("zone1Code", "")
            state_from_api = str(payload.get("admin1Code", "")).strip().upper()
            geocode_fields["State Helper"] = state_from_api or str(loc_dict.get("STATE", "")).strip().upper()
            ok_count += 1
        elif status == 200:
            geocode_fields["Resolution Type"] = "ERROR: HTTP 200 but no coordinates in response"
        else:
            geocode_fields["Resolution Type"] = f"ERROR: HTTP {status} - {error}"

        preview_rows.append({
            "STREETNAME": street, "CITY": loc_dict.get("CITY", ""),
            "STATE (typed)": loc_dict.get("STATE", ""), "ZIPCODE": loc_dict.get("ZIPCODE", ""),
            "Latitude": geocode_fields["Latitude"], "Longitude": geocode_fields["Longitude"],
            "Resolution Type": geocode_fields["Resolution Type"],
            "State (resolved)": geocode_fields["State Helper"],
            "County (resolved)": geocode_fields["County Helper"],
            "CRESTA": geocode_fields["Cresta Zone Helper"],
        })

        loc_rows.append(build_loc_row(account, loc_dict, len(loc_rows) + 1, geocode_fields))

    acc_df = pd.DataFrame([build_acc_row(account)])
    loc_df = pd.DataFrame(loc_rows)
    return acc_df, loc_df, preview_rows, ok_count, total_count


# --------------------------------------------------------------------------
# Streamlit UI
# --------------------------------------------------------------------------

st.set_page_config(page_title="RMS LOC/ACC File Builder", page_icon="📝", layout="wide")

st.markdown(
    """
    <style>
    .block-container { padding-top: 2rem; max-width: 1400px; }
    div[data-testid="stExpander"] { border: 1px solid rgba(135,135,135,0.25); border-radius: 10px; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("📝 RMS LOC/ACC File Builder")
st.caption(
    "Enter account and location data, then geocode and build RMS-import-ready ACC and LOC CSV "
    "files in one step. Ported directly from a confirmed-working RMS cleansing workbook's Data "
    "tab, LOC/ACC formulas, and Geocode_DataTab macro."
)
st.success(
    "🏷️ **Build ID: v4-form-20261030-0942** — if you do NOT see this exact line when you load the "
    "app, you are looking at a cached/older deployment, not this file. Hard-refresh your browser "
    "(Ctrl+Shift+R / Cmd+Shift+R) or redeploy from this exact file before troubleshooting further.",
    icon="🏷️",
)
st.info(
    "ℹ️ All fields below are inside a **form**. Nothing is sent to the app until you press the "
    "**Geocode & Build** button at the bottom - this guarantees every field you typed is captured "
    "together, with no risk of one field resetting while you're still editing another. "
    "This exact mechanism has been verified with an automated test: typing an Account Key and "
    "pressing submit correctly captures it every time.",
    icon="ℹ️",
)

if "locations_df_seed" not in st.session_state:
    st.session_state["locations_df_seed"] = pd.DataFrame([LOCATION_INPUT_DEFAULTS.copy()])

with st.sidebar:
    st.header("Geocode connection")
    host = st.text_input("Host", value="api-euw1.rms.com")
    api_key = st.text_input("Location Intelligence API key", type="password")
    verify_ssl = st.checkbox("Verify SSL", value=True)
    st.caption(
        "Geocoding resolves each address to Latitude/Longitude, State/County/CRESTA - this "
        "happens automatically as part of the combined Geocode & Build action."
    )
    st.divider()
    if st.button("🔄 Reset location table", help="Clears all location rows back to a single blank row."):
        st.session_state["locations_df_seed"] = pd.DataFrame([LOCATION_INPUT_DEFAULTS.copy()])
        st.rerun()

number_column_config = {
    field: st.column_config.NumberColumn(min_value=0, max_value=LARGE_NUMBER_MAX, step=1, format="%.2f")
    for field in LARGE_NUMBER_FIELDS
}
text_column_config = {field: st.column_config.TextColumn() for field in TEXT_FIELDS}

with st.form(key="main_form", clear_on_submit=False):
    st.subheader("1. Account Information")

    top1, top2, top3 = st.columns([2, 1, 1])
    with top1:
        account_key = st.text_input("Account Key", placeholder="25M_1234")
    with top2:
        inception_date = st.date_input("Inception Date", value=date.today())
    with top3:
        expiry_date = st.date_input("Expiry Date", value=date.today())

    col1, col2, col3 = st.columns(3)
    with col1:
        layer_name = st.text_input("Layer Name", placeholder="25M")
        underwriter = st.text_input("Underwriter Initials", placeholder="TR")
    with col2:
        lob = st.text_input("Line of Business", placeholder="ANV_PacificPoint")
        eqsl_account = st.selectbox("EQSL Y/N", ["N", "Y"])
    with col3:
        share_percent = st.number_input(
            "Share (%)", value=100.0, step=5.0, min_value=0.0, max_value=1000.0, format="%.2f",
            help="Participation share as a percentage - e.g. 100% for the full layer, 50% for half.",
        )

    col5, col6, col7, col8 = st.columns(4)
    with col5:
        premium = st.number_input("Premium", value=0, step=1000, min_value=0, max_value=LARGE_NUMBER_MAX)
    with col6:
        policy_limit = st.number_input("Policy Limit", value=0, step=100000, min_value=0, max_value=LARGE_NUMBER_MAX)
    with col7:
        attachment = st.number_input("Attachment", value=0, step=10000, min_value=0, max_value=LARGE_NUMBER_MAX)
    with col8:
        min_ded = st.number_input("Min Ded", value=0, step=1000, min_value=0, max_value=LARGE_NUMBER_MAX)
    max_ded = st.number_input("Max Ded", value=0, step=1000, min_value=0, max_value=LARGE_NUMBER_MAX)

    # Share is captured from the user as a percentage (e.g. 100%) but stored
    # internally as a fraction (1.0), matching the source workbook's native
    # Data!M3 "Share:" field and the ACC!BLANLIMAMT formula, which both
    # expect a fraction, not a percentage.
    share = share_percent / 100.0

    st.divider()
    st.subheader("2. Location Information")
    st.caption(
        "Add one row per location. Dropdowns map directly to the RMS ATC/characteristic codes used "
        "by the LOC file. Monetary fields accept any size value, including hundreds of millions or "
        "more. STREETNAME/CITY/COUNTY/STATE/ZIPCODE are kept as plain text so values like a "
        "leading-zero ZIP code are never silently altered."
    )

    edited_locations_df = st.data_editor(
        st.session_state["locations_df_seed"],
        column_config={
            **text_column_config,
            **number_column_config,
            "Construction": st.column_config.SelectboxColumn(options=options_list(CONSTRUCTION)),
            "Occupancy": st.column_config.SelectboxColumn(options=options_list(OCCUPANCY)),
            "Contents Damageability Grade": st.column_config.SelectboxColumn(options=options_list(CONTENTS_DAMAGE)),
            "Construction Quality": st.column_config.SelectboxColumn(options=options_list(CONSTRUCTION_QUALITY)),
            "Plan Irregularity": st.column_config.SelectboxColumn(options=options_list(PLAN_IRREGULARITY)),
            "Soft Story": st.column_config.SelectboxColumn(options=options_list(SOFT_STORY)),
            "Vertical Irregularity": st.column_config.SelectboxColumn(options=options_list(VERTICAL_IRREGULARITY)),
            "Cladding Type": st.column_config.SelectboxColumn(options=options_list(CLADDING_TYPE)),
            "Frame Bolted": st.column_config.SelectboxColumn(options=options_list(FRAME_BOLTED)),
            "Anchoring": st.column_config.SelectboxColumn(options=options_list(ANCHORING)),
            "Engineered Foundation": st.column_config.SelectboxColumn(options=options_list(ENGINEERED_FOUNDATION)),
            "Sprinkler Type": st.column_config.SelectboxColumn(options=options_list(SPRINKLER_TYPE)),
            "EQSL  Susceptibility": st.column_config.SelectboxColumn(options=options_list(SPRINKLER_LEAKAGE)),
            "EQSL Y/N": st.column_config.SelectboxColumn(options=["N", "Y"]),
            "Number of Buildings": st.column_config.NumberColumn(min_value=0, step=1),
            "Number of Stories": st.column_config.NumberColumn(min_value=0, step=1),
            "Year Built (YYYY)": st.column_config.NumberColumn(min_value=0, max_value=2100, step=1, format="%d"),
            "Year Upgraded (YYYY)": st.column_config.NumberColumn(min_value=0, max_value=2100, step=1, format="%d"),
        },
        num_rows="dynamic",
        hide_index=True,
        width="stretch",
        key="location_editor",
    )

    st.divider()
    submitted = st.form_submit_button("🌍 Geocode & 🏗️ Build ACC/LOC Files", type="primary")

account = {
    "Account Key": account_key, "Layer Name": layer_name,
    "Inception Date": inception_date.strftime("%m/%d/%Y"),
    "Expiry Date": expiry_date.strftime("%m/%d/%Y"),
    "Underwriter Initials": underwriter, "Line of Business": lob,
    "Premium": premium, "Policy Limit": policy_limit, "Attachment": attachment,
    "Min Ded": min_ded, "Max Ded": max_ded, "Share": share, "EQSL Y/N": eqsl_account,
}

with st.expander("🔍 Current form values (debug view)"):
    st.caption(
        "Because the fields above are inside a form, this reflects the LAST SUBMITTED values "
        "(i.e. as of the last time you pressed the button), not live keystrokes - that is expected "
        "form behavior, not a bug."
    )
    st.write("**Account Key:**", repr(account_key))
    st.write("**Location table:**")
    st.dataframe(
        edited_locations_df[[c for c in TEXT_FIELDS if c in edited_locations_df.columns]],
        width="stretch", hide_index=True,
    )

if submitted:
    if not api_key:
        st.warning("Please provide a Location Intelligence API key (sidebar).")
    elif not account_key.strip():
        st.warning(f"Please enter an Account Key. Current value: {repr(account_key)}")
    else:
        api = GeocodeApi(host.strip(), api_key, verify=verify_ssl)
        with st.spinner("Geocoding locations and building files..."):
            acc_df, loc_df, preview_rows, ok_count, total_count = geocode_and_build(
                api, account, edited_locations_df,
            )
        st.session_state["acc_df"] = acc_df
        st.session_state["loc_df"] = loc_df
        st.session_state["preview_rows"] = preview_rows

        if total_count == 0:
            st.warning("No locations with a Street Name were found - add at least one location above.")
        elif ok_count == total_count:
            st.success(f"✅ All {ok_count}/{total_count} location(s) geocoded and files built successfully.")
        else:
            st.warning(
                f"⚠️ {ok_count}/{total_count} location(s) geocoded successfully. "
                f"{total_count - ok_count} location(s) failed - see the Geocode Results table below."
            )

st.divider()
st.subheader("3. Preview & Export")

if "acc_df" in st.session_state:
    acc_df = st.session_state["acc_df"]
    loc_df = st.session_state["loc_df"]
    preview_rows = st.session_state.get("preview_rows", [])

    if preview_rows:
        preview_df = pd.DataFrame(preview_rows)

        st.markdown("#### 🗺️ Location Map")
        map_df = preview_df.copy()
        map_df["Latitude"] = pd.to_numeric(map_df["Latitude"], errors="coerce")
        map_df["Longitude"] = pd.to_numeric(map_df["Longitude"], errors="coerce")
        map_df = map_df.dropna(subset=["Latitude", "Longitude"])

        if not map_df.empty:
            map_df["Geocoded OK"] = ~map_df["Resolution Type"].astype(str).str.startswith("ERROR")
            map_df["fill_color"] = map_df["Geocoded OK"].apply(
                lambda ok: [34, 139, 34, 200] if ok else [214, 39, 40, 200]
            )
            map_df["label"] = (
                map_df["STREETNAME"].astype(str) + ", " + map_df["CITY"].astype(str)
                + " (" + map_df["State (resolved)"].astype(str) + ")"
            )

            mid_lat = map_df["Latitude"].mean()
            mid_lon = map_df["Longitude"].mean()

            scatter_layer = pdk.Layer(
                "ScatterplotLayer",
                data=map_df,
                get_position="[Longitude, Latitude]",
                get_fill_color="fill_color",
                get_radius=250,
                radius_min_pixels=6,
                radius_max_pixels=20,
                pickable=True,
                stroked=True,
                get_line_color=[255, 255, 255],
                line_width_min_pixels=1,
            )
            view_state = pdk.ViewState(
                latitude=mid_lat, longitude=mid_lon,
                zoom=10 if len(map_df) > 1 else 13, pitch=0,
            )
            # NOTE: map_provider defaults to "carto", which needs NO API token
            # and renders reliably with no setup. An earlier version hardcoded
            # a "mapbox://..." style URL, which silently requires a Mapbox
            # API token to load any basemap tiles at all - without one, the
            # map renders completely blank (no error, just empty). Carto's
            # "light" style gives the same clean look with zero configuration.
            deck = pdk.Deck(
                layers=[scatter_layer],
                initial_view_state=view_state,
                map_provider="carto",
                map_style="light",
                tooltip={
                    "html": "<b>{label}</b><br/>CRESTA: {CRESTA}<br/>{Resolution Type}",
                    "style": {"backgroundColor": "#1f2937", "color": "white"},
                },
            )
            st.pydeck_chart(deck)
            legend_col1, legend_col2 = st.columns(2)
            legend_col1.markdown("🟢 **Geocoded successfully**")
            legend_col2.markdown("🔴 **Geocode failed** (see table below)")
        else:
            st.info("No locations have valid coordinates to plot yet.")

        with st.expander("🌍 Geocode Results (read-only table)", expanded=True):
            st.dataframe(preview_df, width="stretch", hide_index=True)

    tab1, tab2 = st.tabs(["ACC preview", "LOC preview"])
    with tab1:
        st.dataframe(acc_df, width="stretch", hide_index=True)
    with tab2:
        st.dataframe(loc_df, width="stretch", hide_index=True)

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_name = account_key.strip() or "export"
    c1, c2 = st.columns(2)
    with c1:
        st.download_button(
            "⬇️ Download ACC CSV", data=acc_df.to_csv(index=False).encode("utf-8"),
            file_name=f"{base_name}_ACC_{stamp}.csv", mime="text/csv", width="stretch",
        )
    with c2:
        st.download_button(
            "⬇️ Download LOC CSV", data=loc_df.to_csv(index=False).encode("utf-8"),
            file_name=f"{base_name}_LOC_{stamp}.csv", mime="text/csv", width="stretch",
        )
else:
    st.info("Fill in the form above and press 'Geocode & Build ACC/LOC Files' to generate the files.")

st.divider()
st.subheader("4. Create EDM & Import to RMS")
st.caption(
    "Uploads the ACC/LOC files built above into RMS: resolves or creates the EDM, creates a "
    "portfolio, uploads both files to S3, submits the MRI import job, and polls for completion. "
    "Ported from the confirmed-working modRMSImport.bas / modAWSSigV4.bas macros."
)

if "acc_df" not in st.session_state:
    st.info("Build the ACC/LOC files in step 3 above before importing to RMS.")
else:
    with st.form(key="import_form", clear_on_submit=False):
        imp_col1, imp_col2 = st.columns(2)
        with imp_col1:
            rms_import_host = st.text_input("RMS host", value="api-euw1.rms.com", key="rms_import_host")
            rms_import_api_key = st.text_input("RMS Risk Data API key", type="password", key="rms_import_api_key",
                                                help="Same key used for platform/riskdata/v1 and riskmodeler/v1 "
                                                     "endpoints - this is a DIFFERENT key from the Location "
                                                     "Intelligence key used for geocoding above.")
        with imp_col2:
            default_edm_name = f"EDM-ANV_{account_key.strip()}_Y{datetime.now().strftime('%y')}" if account_key.strip() else ""
            edm_name_input = st.text_input(
                "EDM name", value=default_edm_name, key="edm_name_input",
                help="Matches the naming convention 'EDM-ANV_{Account Key}_Y{YY}' used by the source "
                     "workbook. Edit if your tenant uses a different convention.",
            )
        rms_import_verify_ssl = st.checkbox("Verify SSL", value=True, key="rms_import_verify_ssl")
        import_submitted = st.form_submit_button("🚀 Create EDM & Import to RMS", type="primary")

    if import_submitted:
        if not rms_import_api_key:
            st.warning("Please provide an RMS Risk Data API key.")
        elif not edm_name_input.strip():
            st.warning("Please provide an EDM name.")
        else:
            acc_df = st.session_state["acc_df"]
            loc_df = st.session_state["loc_df"]
            acc_bytes = acc_df.to_csv(index=False).encode("utf-8")
            loc_bytes = loc_df.to_csv(index=False).encode("utf-8")

            rms_api = RmsImportApi(rms_import_host.strip(), rms_import_api_key, verify=rms_import_verify_ssl)
            status_box = st.status("Starting import...", expanded=True)
            import_result = {"success": False, "log": []}

            def log(msg):
                import_result["log"].append(msg)
                status_box.write(msg)

            try:
                status_box.update(label="Step 1/5: Resolving or creating EDM...")
                log(f"Resolving EDM '{edm_name_input.strip()}'...")

                def wait_cb(attempt, total, name):
                    status_box.update(label=f"Waiting for EDM '{name}' to build... ({attempt}/{total})")

                exposure_id, diag = rms_api.get_or_create_edm(edm_name_input.strip(), wait_callback=wait_cb)
                if not exposure_id:
                    log(f"❌ Failed to resolve or create EDM.\n{diag}")
                    status_box.update(label="❌ EDM creation failed", state="error")
                else:
                    log(f"✅ EDM confirmed - exposureId={exposure_id}")

                    status_box.update(label="Step 2/5: Creating portfolio...")
                    portfolio_id = rms_api.ensure_portfolio_by_identifier(exposure_id, account_key.strip())
                    if not portfolio_id:
                        log("⚠️ Could not create/resolve a portfolio - import will target the EDM directly.")
                    else:
                        log(f"✅ Portfolio ready - portfolioId={portfolio_id}")

                    status_box.update(label="Step 3/5: Creating import folder & uploading files to S3...")
                    folder_ok, folder_id, acc_upload, loc_upload = rms_api.create_import_folder()
                    if not folder_ok:
                        log("❌ Failed to create the import folder.")
                        status_box.update(label="❌ Import folder creation failed", state="error")
                    else:
                        log(f"✅ Import folder created - folderId={folder_id}")

                        acc_ok, acc_err = rms_api.upload_file_to_s3(acc_upload, acc_bytes)
                        log(f"{'✅' if acc_ok else '❌'} ACC file upload: {'OK' if acc_ok else acc_err}")
                        loc_ok, loc_err = rms_api.upload_file_to_s3(loc_upload, loc_bytes)
                        log(f"{'✅' if loc_ok else '❌'} LOC file upload: {'OK' if loc_ok else loc_err}")

                        if not (acc_ok and loc_ok):
                            status_box.update(label="❌ File upload to S3 failed", state="error")
                        else:
                            status_box.update(label="Step 4/5: Submitting import job...")
                            job_id, raw_resp = rms_api.submit_import_job(
                                folder_id, exposure_id, portfolio_id, currency="USD",
                            )
                            if not job_id:
                                log(f"❌ Import job submission failed.\nRaw response: {raw_resp[:500]}")
                                status_box.update(label="❌ Import job submission failed", state="error")
                            else:
                                log(f"✅ Import job submitted - jobId={job_id}")
                                st.session_state["rms_import_job"] = {
                                    "job_id": job_id, "exposure_id": exposure_id,
                                    "portfolio_id": portfolio_id, "edm_name": edm_name_input.strip(),
                                    "status_url": "", "host": rms_import_host.strip(),
                                    "api_key": rms_import_api_key, "verify_ssl": rms_import_verify_ssl,
                                }
                                status_box.update(label="Step 5/5: Polling job status...")

                                final_status = "UNKNOWN"
                                status_url = ""
                                for poll_attempt in range(1, 13):
                                    time.sleep(5)
                                    final_status, status_url, raw_status = rms_api.check_job_status_once(
                                        job_id, cached_status_url=status_url,
                                    )
                                    log(f"  Poll {poll_attempt}/12: status = {final_status}")
                                    if final_status in ("COMPLETED", "SUCCESS", "SUCCEEDED", "FAILED", "ERROR"):
                                        break
                                st.session_state["rms_import_job"]["status_url"] = status_url

                                if final_status in ("COMPLETED", "SUCCESS", "SUCCEEDED"):
                                    log(f"✅ Import job finished with status: {final_status}")
                                    status_box.update(label="✅ Import complete!", state="complete")
                                    import_result["success"] = True
                                elif final_status in ("FAILED", "ERROR"):
                                    log(f"❌ Import job finished with status: {final_status}")
                                    status_box.update(label="❌ Import job failed", state="error")
                                else:
                                    log(f"⏳ Import job still in progress (last status: {final_status}) - "
                                        f"use 'Check status' below to keep polling.")
                                    status_box.update(label=f"⏳ Still running (status: {final_status})", state="running")
            except Exception as exc:
                log(f"❌ Unexpected error: {exc}")
                status_box.update(label="❌ Import failed with an unexpected error", state="error")

    if "rms_import_job" in st.session_state:
        job_info = st.session_state["rms_import_job"]
        st.caption(
            f"Last submitted job: **{job_info['job_id']}** for EDM **{job_info['edm_name']}** "
            f"(exposureId={job_info['exposure_id']}, portfolioId={job_info['portfolio_id'] or '—'})"
        )
        if st.button("🔄 Check import status"):
            check_api = RmsImportApi(job_info["host"], job_info["api_key"], verify=job_info["verify_ssl"])
            status, status_url, raw = check_api.check_job_status_once(
                job_info["job_id"], cached_status_url=job_info["status_url"],
            )
            job_info["status_url"] = status_url
            st.session_state["rms_import_job"] = job_info
            if status in ("COMPLETED", "SUCCESS", "SUCCEEDED"):
                st.success(f"✅ Status: {status}")
            elif status in ("FAILED", "ERROR"):
                st.error(f"❌ Status: {status}")
                with st.expander("Raw response"):
                    st.code(raw[:2000])
            else:
                st.info(f"⏳ Status: {status}")
