#!/usr/bin/env python
"""
RMS LOC/ACC File Builder (Streamlit) - v2

Ports the "Data" tab + its live formulas (LOC/ACC sheets) + Geocode_DataTab.bas
macro from a confirmed-working RMS cleansing workbook into a single Streamlit
data-entry app. Produces RMS-import-ready ACC and LOC CSV files.

Workflow:
  1. Enter Account Information (one record - mirrors Data!A2:N3).
  2. Enter one or more Locations (mirrors Data!B11:AV11 input columns).
  3. Click "Geocode & Build ACC/LOC Files" - a SINGLE combined action that
     geocodes every location and builds the ACC/LOC tables in one pass.

ARCHITECTURE NOTE - v2 fix for a real bug found in v1
------------------------------------------------------
v1 tried to write geocode results (Latitude, Longitude, State/County/CRESTA)
back INTO the same st.data_editor used for location input, across a
separate button click and page rerun. This relies on Streamlit correctly
re-adopting a changed `data` argument for an existing widget `key` -  which,
in practice (especially with num_rows="dynamic" editors), does not reliably
sync on every Streamlit version, and/or the "geocoded OK" counter was only
checking the HTTP status rather than confirming usable coordinates were
actually present in the response. Both of these could silently leave
Latitude/Longitude blank even after a reported "Geocoded 1/1" success.

v2 removes the round-trip entirely:
  - The location editor ONLY ever holds the fields the user types. It never
    displays or stores geocode outputs.
  - Geocoding and file-building happen in ONE button handler, in ONE script
    run, operating on the data_editor's just-returned value as a plain
    local Python variable (not session_state) - so there is no cross-run
    synchronization to get wrong.
  - A row only counts as successfully geocoded if the response contains a
    non-null, non-zero latitude AND longitude - not merely HTTP 200.
  - Geocode results are shown in a separate, read-only preview table built
    fresh every time the combined action runs.

Known issues in the SOURCE WORKBOOK this app intentionally does NOT
reproduce (fixed here instead):
  - LOC!T2 (YEARBUILT) and LOC!AS2 (YEARUPGRAD) build a string
    "01/01/"&year. RMS needs these as PLAIN 4-DIGIT YEAR INTEGERS, not
    dates - this app outputs the bare year.
  - ACC!Y2 (BLANDEDCUR) contains a broken #REF! formula from a deleted
    column - this app outputs an empty value for that field instead.
  - LOC Latitude/Longitude/ADDRMATCH/Resolution Type columns have no
    formula at all in the live sheet - this app populates them from the
    geocode step, which is clearly the columns' intended purpose.
"""

import io
import time
from datetime import date, datetime

import pandas as pd
import requests
import streamlit as st
from requests.adapters import HTTPAdapter

TIMEOUT = 60
MAX_RETRIES = 4
LARGE_NUMBER_MAX = 999_999_999_999  # effectively unbounded for insurance values (up to ~$1 trillion)

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
# NOTE: geocode OUTPUT fields (Latitude, Longitude, Resolution Code/Type,
# State/County/CRESTA helpers) are intentionally NOT included here anymore -
# they are computed fresh every time the combined Geocode & Build action
# runs, and shown in a separate read-only preview table. This is the core
# v2 fix: the editable table never needs to "receive back" computed data.
# --------------------------------------------------------------------------

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

# Monetary / large-number fields that must accept big values without
# truncation or silent clamping - addressed explicitly via NumberColumn
# configs below (min_value=0, no max_value cap except a generous ceiling,
# and a format that does not round or display in scientific notation).
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
    this location; `geocode_fields` is a dict with Latitude, Longitude,
    Resolution Code, Resolution Type, State Helper, County Helper, Cresta
    Zone Helper - computed fresh in the SAME pass, never round-tripped
    through the editor."""
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
    builds the matching LOC row immediately, using only local variables.
    Returns (acc_df, loc_df, preview_rows, ok_count, total_count)."""
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

        # A row only counts as successfully geocoded if BOTH coordinates
        # are present and non-null - matching HTTP 200 alone is NOT enough
        # (this was the root cause of v1's "Geocoded 1/1 but still blank"
        # bug: it counted success on status code only).
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

if "locations_df" not in st.session_state:
    st.session_state["locations_df"] = pd.DataFrame([LOCATION_INPUT_DEFAULTS.copy()])

with st.sidebar:
    st.header("Geocode connection")
    host = st.text_input("Host", value="api-euw1.rms.com")
    api_key = st.text_input("Location Intelligence API key", type="password")
    verify_ssl = st.checkbox("Verify SSL", value=True)
    st.caption(
        "Geocoding resolves each address to Latitude/Longitude, State/County/CRESTA - this "
        "happens automatically as part of the combined Geocode & Build action below."
    )

st.subheader("1. Account Information")
col1, col2, col3, col4 = st.columns(4)
with col1:
    account_key = st.text_input("Account Key", placeholder="25M_1234")
    layer_name = st.text_input("Layer Name", placeholder="25M")
with col2:
    inception_date = st.date_input("Inception Date", value=date.today())
    expiry_date = st.date_input("Expiry Date", value=date.today())
with col3:
    underwriter = st.text_input("Underwriter Initials", placeholder="TR")
    lob = st.text_input("Line of Business", placeholder="ANV_PacificPoint")
with col4:
    eqsl_account = st.selectbox("EQSL Y/N", ["N", "Y"])

col5, col6, col7, col8, col9 = st.columns(5)
with col5:
    premium = st.number_input("Premium", value=0, step=1000, min_value=0, max_value=LARGE_NUMBER_MAX)
with col6:
    policy_limit = st.number_input("Policy Limit", value=0, step=100000, min_value=0, max_value=LARGE_NUMBER_MAX)
with col7:
    attachment = st.number_input("Attachment", value=0, step=10000, min_value=0, max_value=LARGE_NUMBER_MAX)
with col8:
    min_ded = st.number_input("Min Ded", value=0, step=1000, min_value=0, max_value=LARGE_NUMBER_MAX)
with col9:
    max_ded = st.number_input("Max Ded", value=0, step=1000, min_value=0, max_value=LARGE_NUMBER_MAX)
share = st.number_input("Share", value=1.0, step=0.05, format="%.4f", min_value=0.0)

account = {
    "Account Key": account_key, "Layer Name": layer_name,
    "Inception Date": inception_date.strftime("%m/%d/%Y"),
    "Expiry Date": expiry_date.strftime("%m/%d/%Y"),
    "Underwriter Initials": underwriter, "Line of Business": lob,
    "Premium": premium, "Policy Limit": policy_limit, "Attachment": attachment,
    "Min Ded": min_ded, "Max Ded": max_ded, "Share": share, "EQSL Y/N": eqsl_account,
}

st.divider()
st.subheader("2. Location Information")
st.caption(
    "Add one row per location. Dropdowns map directly to the RMS ATC/characteristic codes used by "
    "the LOC file. Monetary fields (Building Value, Contents, etc.) accept any size value, "
    "including figures in the hundreds of millions or more."
)

number_column_config = {
    field: st.column_config.NumberColumn(min_value=0, max_value=LARGE_NUMBER_MAX, step=1, format="%.2f")
    for field in LARGE_NUMBER_FIELDS
}

locations_df = st.data_editor(
    st.session_state["locations_df"],
    column_config={
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
        "STATE": st.column_config.TextColumn(help="2-letter code (CA) or full name (California) - either works."),
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
st.session_state["locations_df"] = locations_df

st.divider()
st.subheader("3. Geocode & Build")
st.caption(
    "This single action geocodes every location above and builds the ACC/LOC files in one pass - "
    "no separate steps required, so results are always in sync with what you see in the table."
)
build_btn = st.button("🌍 Geocode & 🏗️ Build ACC/LOC Files", type="primary")

if build_btn:
    if not api_key:
        st.warning("Please provide a Location Intelligence API key (sidebar).")
    elif not account_key.strip():
        st.warning("Please enter an Account Key.")
    else:
        api = GeocodeApi(host.strip(), api_key, verify=verify_ssl)
        with st.spinner("Geocoding locations and building files..."):
            acc_df, loc_df, preview_rows, ok_count, total_count = geocode_and_build(
                api, account, locations_df,
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
                f"{total_count - ok_count} location(s) failed - see the Geocode Results table below "
                f"for details on each failure."
            )

st.divider()
st.subheader("4. Preview & Export")

if "acc_df" in st.session_state:
    acc_df = st.session_state["acc_df"]
    loc_df = st.session_state["loc_df"]
    preview_rows = st.session_state.get("preview_rows", [])

    if preview_rows:
        with st.expander("🌍 Geocode Results (read-only)", expanded=True):
            st.dataframe(pd.DataFrame(preview_rows), width="stretch", hide_index=True)

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
    st.info("Run step 3 to generate the ACC/LOC files.")
