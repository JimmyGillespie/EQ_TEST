#!/usr/bin/env python
"""
RMS Locations + Hazard Enrichment Explorer (Streamlit)

Ports modLocationsHazard.bas (a confirmed-working RMS VBA macro) into a
Streamlit app. Workflow, per Exposure ID:

  STAGE 1 - Pull locations
    GET /platform/riskdata/v1/exposures/{exposureId}/locations  (paginated)
    -> filter to the highest accountId present (the most recent submission)
    GET .../locations/{locationId}                              (per location)
    -> building TIV breakdown, EQSLINS flag, characteristics

  STAGE 2 - Hazard enrichment (per US location, concurrent across locations)
    POST /li/na_eq_hazard/latest   -> soil / liquefaction / landslide
    POST /li/us_mmi/latest         -> earthquake shaking intensity (MMI)
    POST /li/us_fl_depth/latest    -> flood depth (100/500-yr, undef/def)
    POST /li/us_fema_firm/latest   -> FEMA flood zone
    POST /li/us_fl_risk_score/latest -> flood risk score (needs building chars)
    POST /li/geocode/latest        -> CRESTA zone (zone1Code)

Two SEPARATE API keys are required, matching the working macro exactly:
  - Risk Data API key   (platform/riskdata/v1 - locations/accounts)
  - Hazard/LI API key   (li/* - Location Intelligence hazard layers)

Admin1 (state) resolution fallback chain mirrors the macro's:
  1. Use admin1Code/admin1Name directly from the location address if present.
  2. Optional "Location state fallback" CSV upload (AccountNumber,
     LocationNumber, StateCode, PostalCode) - equivalent of the macro's LOC
     sheet lookup.
  3. Optional "ZIP to state" CSV upload (zip, state_id) - equivalent of the
     macro's uszips sheet lookup.
  US hazard layers are only run once a location has an admin1Code successfully
  resolved - same gate the macro applies before firing any hazard calls.
"""

import io
import json
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
import requests
import streamlit as st
from requests.adapters import HTTPAdapter

# --------------------------------------------------------------------------
# Config
# --------------------------------------------------------------------------

TIMEOUT = 60
MAX_RETRIES = 4
PAGE_LIMIT = 1000

US_STATE_NAME_TO_CODE = {
    "ALABAMA": "AL", "ALASKA": "AK", "ARIZONA": "AZ", "ARKANSAS": "AR", "CALIFORNIA": "CA",
    "COLORADO": "CO", "CONNECTICUT": "CT", "DELAWARE": "DE", "FLORIDA": "FL", "GEORGIA": "GA",
    "HAWAII": "HI", "IDAHO": "ID", "ILLINOIS": "IL", "INDIANA": "IN", "IOWA": "IA",
    "KANSAS": "KS", "KENTUCKY": "KY", "LOUISIANA": "LA", "MAINE": "ME", "MARYLAND": "MD",
    "MASSACHUSETTS": "MA", "MICHIGAN": "MI", "MINNESOTA": "MN", "MISSISSIPPI": "MS", "MISSOURI": "MO",
    "MONTANA": "MT", "NEBRASKA": "NE", "NEVADA": "NV", "NEW HAMPSHIRE": "NH", "NEW JERSEY": "NJ",
    "NEW MEXICO": "NM", "NEW YORK": "NY", "NORTH CAROLINA": "NC", "NORTH DAKOTA": "ND", "OHIO": "OH",
    "OKLAHOMA": "OK", "OREGON": "OR", "PENNSYLVANIA": "PA", "RHODE ISLAND": "RI", "SOUTH CAROLINA": "SC",
    "SOUTH DAKOTA": "SD", "TENNESSEE": "TN", "TEXAS": "TX", "UTAH": "UT", "VERMONT": "VT",
    "VIRGINIA": "VA", "WASHINGTON": "WA", "WEST VIRGINIA": "WV", "WISCONSIN": "WI", "WYOMING": "WY",
    "DISTRICT OF COLUMBIA": "DC",
}
VALID_STATE_CODES = set(US_STATE_NAME_TO_CODE.values())

MMI_LABELS = {
    1: "Not Felt", 2: "Weak", 3: "Weak", 4: "Light", 5: "Moderate", 6: "Strong",
    7: "Very Strong", 8: "Severe", 9: "Violent", 10: "Extreme", 11: "Extreme", 12: "Extreme",
}

# 37-column header list, in exact order, ported from LH_HeaderList() in the
# confirmed-working VBA macro.
LOCATION_CORE_HEADERS = [
    "LocID", "Location Name", "Account Name", "Account Number",
    "Location TIV", "Buildings TIV", "Contents TIV", "BI TIV", "EQSLINS",
    "Construction Name", "Construction Scheme", "Construction Code",
    "Occupancy Name", "Occupancy Scheme", "Occupancy Code",
    "# of Stories", "Year Built", "User Text 1", "User Text 2", "User Defined 1", "User Defined 2",
    "Country", "CRESTA", "Admin1", "Admin2", "Admin3", "City", "Admin4", "Postal Code", "Street Address",
    "Latitude", "Longitude", "Geocoding Resolution", "Perils",
    "Year Built Band", "Stories Band", "Lat,Long",
]

HAZARD_HEADERS = [
    "floodZone", "floodDepthUndef100yr", "floodDepthDef100yr", "floodDepthUndef500yr", "floodDepthDef500yr",
    "floodScoreOverall", "floodScore100yr", "floodScore250yr", "floodScore500yr",
    "soilDescription", "liquefactionDescription", "landslideDescription", "mmi100yr", "MMI_Description",
]


# --------------------------------------------------------------------------
# Risk Data API client (platform/riskdata/v1) - pulls locations + detail
# --------------------------------------------------------------------------

class RiskDataApi:
    """Wraps GET .../exposures/{id}/locations and .../locations/{locId},
    exactly as called in LH_PullLocationsAndHazardsForAccount."""

    def __init__(self, host, api_key, verify=True, pool_size=16):
        self.base_url = f"https://{host}/platform/riskdata/v1"
        self.session = requests.Session()
        adapter = HTTPAdapter(pool_connections=pool_size, pool_maxsize=pool_size)
        self.session.mount("https://", adapter)
        self.session.mount("http://", adapter)
        clean_key = (api_key or "").strip()
        if clean_key.lower().startswith("bearer "):
            clean_key = clean_key[7:].strip()
        self.session.headers.update({"Accept": "application/json", "Authorization": clean_key})
        self.session.verify = verify
        self.lock = threading.Lock()
        self.ssl_fallback_done = not verify

    def _disable_ssl_verify(self):
        with self.lock:
            if self.ssl_fallback_done:
                return
            import urllib3
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
            self.session.verify = False
            self.ssl_fallback_done = True

    def _get(self, url, retries=MAX_RETRIES):
        last_exc = None
        for attempt in range(retries):
            try:
                response = self.session.get(url, timeout=TIMEOUT)
                if response.status_code == 200:
                    try:
                        return response.json(), response.status_code, None
                    except ValueError:
                        return None, response.status_code, "Response was not valid JSON."
                if response.status_code in (429, 500, 502, 503, 504) and attempt < retries - 1:
                    time.sleep(min(2 ** (attempt + 1), 15))
                    continue
                try:
                    detail = response.json()
                except ValueError:
                    detail = response.text[:500]
                return None, response.status_code, detail
            except requests.exceptions.SSLError:
                if self.session.verify:
                    self._disable_ssl_verify()
                    continue
                return None, 0, "SSL error persisted with verification disabled."
            except requests.RequestException as exc:
                last_exc = str(exc)
                if attempt < retries - 1:
                    time.sleep(min(2 ** (attempt + 1), 15))
                    continue
                return None, 0, last_exc
        return None, 0, last_exc

    def list_all_locations(self, exposure_id, progress_cb=None):
        """Paginates through every location in the exposure. Returns
        (list_of_location_wrappers, error_message_or_None)."""
        all_locations = []
        offset = 0
        while True:
            url = f"{self.base_url}/exposures/{exposure_id}/locations?limit={PAGE_LIMIT}&offset={offset}"
            payload, status, error = self._get(url)
            if status != 200 or payload is None:
                if offset == 0:
                    return [], f"HTTP {status}: {error}"
                break  # partial results already gathered - stop paginating
            if not payload:
                break
            all_locations.extend(payload)
            if progress_cb:
                progress_cb(len(all_locations))
            if len(payload) < PAGE_LIMIT:
                break
            offset += PAGE_LIMIT
            if offset > 200000:
                break
        return all_locations, None

    def get_location_detail(self, exposure_id, location_id):
        url = f"{self.base_url}/exposures/{exposure_id}/locations/{location_id}"
        return self._get(url)


# --------------------------------------------------------------------------
# Location Intelligence API client (li/*) - hazard + geocode lookups
# --------------------------------------------------------------------------

class LocationIntelligenceApi:
    """Confirmed-working request shapes, ported from modFloodGeo.bas and
    modLocationsHazard.bas. Three body styles:
      "minimal"     -> {"location":{"address":{...}}}                (EQ, MMI)
      "flood_basic" -> {"location":{"address":{...},"characteristics":{},
                         "coverageValues":{}}, "layerOptions":{...}}  (FEMA, depth)
      "full"        -> {"location":{"address":{...},
                         "characteristics":{...},"coverageValues":{...}}} (risk score)
    """

    def __init__(self, host, api_key, verify=True, pool_size=24):
        self.base_url = f"https://{host}/li"
        self.session = requests.Session()
        adapter = HTTPAdapter(pool_connections=pool_size, pool_maxsize=pool_size)
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
        self.lock = threading.Lock()
        self.ssl_fallback_done = not verify

    def _disable_ssl_verify(self):
        with self.lock:
            if self.ssl_fallback_done:
                return
            import urllib3
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
            self.session.verify = False
            self.ssl_fallback_done = True

    def _post(self, path, body, retries=MAX_RETRIES):
        url = f"{self.base_url}/{path}"
        last_exc = None
        for attempt in range(retries):
            try:
                response = self.session.post(url, json=body, timeout=TIMEOUT)
                if response.status_code == 200:
                    try:
                        return response.json(), response.status_code, None
                    except ValueError:
                        return None, response.status_code, "Response was not valid JSON."
                if response.status_code in (429, 500, 502, 503, 504) and attempt < retries - 1:
                    time.sleep(min(2 ** (attempt + 1), 15))
                    continue
                try:
                    detail = response.json()
                except ValueError:
                    detail = response.text[:500]
                return None, response.status_code, detail
            except requests.exceptions.SSLError:
                if self.session.verify:
                    self._disable_ssl_verify()
                    continue
                return None, 0, "SSL error persisted with verification disabled."
            except requests.RequestException as exc:
                last_exc = str(exc)
                if attempt < retries - 1:
                    time.sleep(min(2 ** (attempt + 1), 15))
                    continue
                return None, 0, last_exc
        return None, 0, last_exc

    def geocode_flat(self, street, city, admin1_code, postal, country_code, version="latest"):
        """Flat body for CRESTA lookup - matches jsonBodyGeocode in the macro."""
        body = {"countryScheme": "ISO2A", "countryCode": country_code}
        if admin1_code:
            body["admin1Code"] = admin1_code
        if city:
            body["cityName"] = city
        if postal:
            body["postalCode"] = postal
        if street:
            body["streetAddress"] = street
        return self._post(f"geocode/{version}", body)

    def hazard_lookup(self, lat, lon, admin1_code, country_code, resolution_code,
                       layer, version="latest", body_style="minimal",
                       layer_options=None, characteristics=None, coverage_values=None):
        address = {
            "countryScheme": "ISO2A", "countryCode": country_code, "countryRmsCode": country_code,
        }
        if admin1_code:
            address["admin1Code"] = admin1_code
        address["latitude"] = lat
        address["longitude"] = lon
        address["rmsGeoModelResolutionCode"] = resolution_code

        location = {"address": address}
        if body_style == "minimal":
            pass
        elif body_style == "flood_basic":
            location["characteristics"] = {}
            location["coverageValues"] = {}
        else:  # "full"
            location["characteristics"] = characteristics or {}
            location["coverageValues"] = coverage_values or {}

        body = {"location": location}
        if body_style == "flood_basic" and layer_options:
            body["layerOptions"] = layer_options

        return self._post(f"{layer}/{version}", body)


# --------------------------------------------------------------------------
# JSON / data helpers (ported from LH_SafeGet, LH_StripPrefix, LH_GetAdmin1,
# LH_GetCresta, LH_IsBlankVal, LH_NormalizeZip)
# --------------------------------------------------------------------------

def safe_get(obj, key, child_key=None):
    if not isinstance(obj, dict):
        return 0
    if child_key is None:
        val = obj.get(key)
        if val is None or val == "":
            return 0
        return val
    child = obj.get(key)
    if not isinstance(child, dict):
        return 0
    val = child.get(child_key)
    if val is None or val == "":
        return 0
    return val


def is_blank_val(v):
    if v is None:
        return True
    if isinstance(v, (int, float)):
        return v == 0
    return len(str(v).strip()) == 0


def strip_prefix(txt):
    """'1053 - Fire Resistive' -> 'Fire Resistive' (ported from LH_StripPrefix)."""
    txt = str(txt or "")
    parts = txt.split(" - ", 1)
    return parts[1] if len(parts) > 1 else txt


def normalize_zip(raw):
    s = str(raw or "").strip()
    if "-" in s:
        s = s.split("-")[0]
    digits = "".join(ch for ch in s if ch.isdigit())
    if not digits:
        return ""
    if len(digits) < 5:
        digits = digits.zfill(5)
    return digits[:5]


def get_admin1_from_address(address):
    """Ported from LH_GetAdmin1: tries admin1Name, admin1.name, stateName, state."""
    for key in ("admin1Name",):
        v = safe_get(address, key)
        if not is_blank_val(v):
            return str(v)
    v = safe_get(address, "admin1", "name")
    if not is_blank_val(v):
        return str(v)
    for key in ("stateName", "state"):
        v = safe_get(address, key)
        if not is_blank_val(v):
            return str(v)
    return ""


def resolve_admin1_code(admin1_text, direct_admin1_code, fallback_df, zip_df,
                         account_number, loc_number, direct_postal_code):
    """Resolves a 2-letter admin1Code for hazard calls. Fallback chain
    (ported from the macro's state dict -> LOC-sheet -> uszips chain):
      1. admin1_text is a full state name we recognize -> map to code.
      2. direct_admin1_code is already a valid 2-letter code -> use it.
      3. Optional fallback_df (AccountNumber, LocationNumber, StateCode,
         PostalCode CSV upload) matched on account+location.
      4. Optional zip_df (zip, state_id CSV upload) matched on postal code,
         trying the location's own postal code first, then the fallback
         CSV's postal code for that account+location.
    Returns ("", "") if nothing resolves.
    """
    name_upper = str(admin1_text or "").strip().upper()
    if name_upper in US_STATE_NAME_TO_CODE:
        code = US_STATE_NAME_TO_CODE[name_upper]
        return code, admin1_text

    code_upper = str(direct_admin1_code or "").strip().upper()
    if len(code_upper) == 2 and code_upper in VALID_STATE_CODES:
        return code_upper, admin1_text or code_upper

    if fallback_df is not None and not fallback_df.empty:
        match = fallback_df[
            (fallback_df["AccountNumber"].astype(str).str.upper() == str(account_number).strip().upper())
            & (fallback_df["LocationNumber"].astype(str) == str(loc_number).strip())
        ]
        if not match.empty:
            fb_code = str(match.iloc[0].get("StateCode", "")).strip().upper()
            if len(fb_code) == 2 and fb_code in VALID_STATE_CODES:
                return fb_code, admin1_text

            fb_zip = normalize_zip(match.iloc[0].get("PostalCode", ""))
            zip_try = normalize_zip(direct_postal_code) or fb_zip
            if zip_try and zip_df is not None and not zip_df.empty:
                zmatch = zip_df[zip_df["zip"].astype(str) == zip_try]
                if not zmatch.empty:
                    z_code = str(zmatch.iloc[0].get("state_id", "")).strip().upper()
                    if len(z_code) == 2:
                        return z_code, admin1_text

    zip_try = normalize_zip(direct_postal_code)
    if zip_try and zip_df is not None and not zip_df.empty:
        zmatch = zip_df[zip_df["zip"].astype(str) == zip_try]
        if not zmatch.empty:
            z_code = str(zmatch.iloc[0].get("state_id", "")).strip().upper()
            if len(z_code) == 2:
                return z_code, admin1_text

    return "", admin1_text


def year_built_band(year_val):
    try:
        y = float(year_val)
    except (TypeError, ValueError):
        return "Unknown"
    if y == 9999:
        return "Unknown"
    if y >= 2020:
        return "2020-2030"
    if y >= 2010:
        return "2010-2020"
    if y >= 2000:
        return "2000-2010"
    if y >= 1990:
        return "1990-2000"
    return "Pre-1990"


def stories_band(stories_val):
    try:
        s = float(stories_val)
    except (TypeError, ValueError):
        return "Unknown"
    if s >= 50:
        return "50+"
    if s >= 21:
        return "21-50"
    if s >= 11:
        return "11-20"
    if s >= 7:
        return "7-10"
    return str(int(s)) if s == int(s) else str(s)


def mmi_label(value):
    try:
        v = round(float(value))
    except (TypeError, ValueError):
        return "Unknown"
    v = max(1, min(12, v))
    return MMI_LABELS.get(v, "Unknown")


# --------------------------------------------------------------------------
# STAGE 1: Pull locations + detail (ported from
# LH_PullLocationsAndHazardsForAccount, detail portion)
# --------------------------------------------------------------------------

def find_max_account_id(all_locations):
    max_acct = -1
    for lw in all_locations:
        location = lw.get("location") or {}
        prop = location.get("property") or {}
        acct_id = safe_get(prop, "accountId")
        try:
            acct_id_f = float(acct_id)
            if acct_id_f > max_acct:
                max_acct = acct_id_f
        except (TypeError, ValueError):
            continue
    return max_acct


def build_core_row(lw, detail_json, account_number):
    """Builds one 37-field row dict, ported field-by-field from the VBA's
    outArr(1..37) construction in LH_PullLocationsAndHazardsForAccount."""
    location = lw.get("location") or {}
    prop = location.get("property") or {}
    address = location.get("address") or {}

    row = {}
    row["LocID"] = safe_get(prop, "locationNumber")
    row["Location Name"] = safe_get(prop, "locationName")
    row["Account Name"] = account_number
    row["Account Number"] = safe_get(prop, "accountId")

    tiv_total = lw.get("tiv", location.get("tiv", 0))
    try:
        row["Location TIV"] = float(tiv_total)
    except (TypeError, ValueError):
        row["Location TIV"] = 0

    tiv_bldg, tiv_contents, tiv_bi, eqsl_ins = 0, 0, 0, ""
    if isinstance(detail_json, dict):
        try:
            tiv_bldg = float(safe_get(detail_json, "maxCvgBldgValuation") or 0)
        except (TypeError, ValueError):
            tiv_bldg = 0
        for cov in (detail_json.get("locationCoverages") or []):
            loss_type = str(cov.get("lossType", "")).lower()
            if loss_type == "contents":
                try:
                    tiv_contents = float(cov.get("value", 0))
                except (TypeError, ValueError):
                    pass
            elif loss_type in ("bi", "business interruption"):
                try:
                    tiv_bi = float(cov.get("value", 0))
                except (TypeError, ValueError):
                    pass
        eq_detail = detail_json.get("locationEqDetail") or {}
        eqsl_obj = eq_detail.get("eqslInsurance")
        if isinstance(eqsl_obj, dict):
            combined = (str(eqsl_obj.get("code", "")).upper() + "|" + str(eqsl_obj.get("name", "")).upper())
            if "YES" in combined:
                eqsl_ins = "1"
            elif "NO" in combined:
                eqsl_ins = "0"

    row["Buildings TIV"] = tiv_bldg
    row["Contents TIV"] = tiv_contents
    row["BI TIV"] = tiv_bi
    row["EQSLINS"] = eqsl_ins

    row["Construction Name"] = strip_prefix(str(safe_get(prop, "buildingClass", "name")))
    row["Construction Scheme"] = safe_get(prop, "buildingClassScheme", "name")
    row["Construction Code"] = safe_get(prop, "buildingClass", "code")

    row["Occupancy Name"] = strip_prefix(str(safe_get(prop, "occupancyType", "name")))
    row["Occupancy Scheme"] = safe_get(prop, "occupancyTypeScheme", "name")
    row["Occupancy Code"] = safe_get(prop, "occupancyType", "id")

    row["# of Stories"] = safe_get(prop, "stories")
    row["Year Built"] = safe_get(prop, "yearBuilt")
    row["User Text 1"] = safe_get(prop, "userText1")
    row["User Text 2"] = safe_get(prop, "userText2")
    row["User Defined 1"] = safe_get(prop, "userId1")
    row["User Defined 2"] = safe_get(prop, "userId2")

    row["Country"] = safe_get(address, "country", "name")
    row["CRESTA"] = ""  # filled in Stage 2 from the geocode lookup (zone1Code)
    row["Admin1"] = get_admin1_from_address(address)
    row["Admin2"] = safe_get(address, "admin2Name")
    row["Admin3"] = safe_get(address, "admin3Name")
    row["City"] = safe_get(address, "cityName")
    row["Admin4"] = safe_get(address, "admin4Name")
    row["Postal Code"] = safe_get(address, "postalCode")
    row["Street Address"] = safe_get(address, "streetAddress")

    row["Latitude"] = safe_get(address, "latitude")
    row["Longitude"] = safe_get(address, "longitude")
    row["Geocoding Resolution"] = safe_get(address, "geoResolutionCode", "name")

    perils = prop.get("perils") or []
    row["Perils"] = ", ".join(str(p) for p in perils) if isinstance(perils, list) else str(perils)

    row["Year Built Band"] = year_built_band(row["Year Built"])
    row["Stories Band"] = stories_band(row["# of Stories"])
    row["Lat,Long"] = f"{row['Latitude']},{row['Longitude']}"

    row["_locationId"] = safe_get(prop, "locationId")
    return row


def pull_locations_for_exposure(api, exposure_id, account_number, max_workers, progress_cb=None):
    """STAGE 1 orchestration. Returns (list_of_row_dicts, error_message_or_None)."""
    all_locations, list_error = api.list_all_locations(exposure_id)
    if list_error:
        return [], f"Locations list failed: {list_error}"
    if not all_locations:
        return [], "No locations returned for this Exposure ID."

    max_acct = find_max_account_id(all_locations)
    current = [lw for lw in all_locations
               if str(safe_get((lw.get("location") or {}).get("property") or {}, "accountId")) == str(int(max_acct))]
    if not current:
        return [], f"No locations matched the highest accountId ({max_acct})."

    rows = [None] * len(current)

    def _fetch_one(i, lw):
        prop = (lw.get("location") or {}).get("property") or {}
        loc_id = safe_get(prop, "locationId")
        detail_json, status, _ = api.get_location_detail(exposure_id, loc_id) if loc_id else (None, 0, None)
        return i, build_core_row(lw, detail_json if status == 200 else None, account_number)

    completed = 0
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [executor.submit(_fetch_one, i, lw) for i, lw in enumerate(current)]
        for future in as_completed(futures):
            i, row = future.result()
            rows[i] = row
            completed += 1
            if progress_cb:
                progress_cb(completed, len(current))

    return rows, None


# --------------------------------------------------------------------------
# STAGE 2: Hazard enrichment per US location (ported from
# LH_AddHazardsForRowRange)
# --------------------------------------------------------------------------

def enrich_row_with_hazards(li_api, row, fallback_df, zip_df, account_number):
    """Runs the 6 confirmed hazard/geocode calls for one location row and
    returns a dict of hazard fields to merge in. Non-US locations, or rows
    that fail admin1 resolution, are skipped exactly as the macro does."""
    out = {h: "" for h in HAZARD_HEADERS}
    out["CRESTA"] = ""

    country = str(row.get("Country", "")).strip().upper()
    if country not in ("US", "UNITED STATES", ""):
        out["_hazard_note"] = "Non-US location - hazard layers not run"
        return out

    lat, lon = row.get("Latitude"), row.get("Longitude")
    try:
        lat_f, lon_f = float(lat), float(lon)
        if lat_f == 0 and lon_f == 0:
            raise ValueError
    except (TypeError, ValueError):
        out["_hazard_note"] = "Not geocoded - no Latitude/Longitude on this location"
        return out

    admin1_code, _ = resolve_admin1_code(
        row.get("Admin1"), row.get("Admin1"), fallback_df, zip_df,
        account_number, row.get("LocID"), row.get("Postal Code"),
    )
    if not admin1_code:
        out["_hazard_note"] = "Could not resolve a valid 2-letter admin1Code - hazard layers not run"
        return out

    resolution_code = 1  # "Coordinate" - matches the macro's hard-coded value

    # --- Earthquake: soil / liquefaction / landslide ---
    payload, status, error = li_api.hazard_lookup(
        lat_f, lon_f, admin1_code, "US", resolution_code, "na_eq_hazard", body_style="minimal")
    if status == 200 and isinstance(payload, dict):
        out["soilDescription"] = payload.get("soilDescription", "N/A")
        out["liquefactionDescription"] = payload.get("liquefactionDescription", "N/A")
        out["landslideDescription"] = payload.get("landslideDescription", "N/A")
    else:
        out["soilDescription"] = out["liquefactionDescription"] = out["landslideDescription"] = "Error"

    # --- MMI (earthquake shaking intensity) ---
    payload, status, error = li_api.hazard_lookup(
        lat_f, lon_f, admin1_code, "US", resolution_code, "us_mmi", body_style="minimal")
    if status == 200 and isinstance(payload, dict) and "mmi100yr" in payload:
        mmi_val = payload.get("mmi100yr")
        out["mmi100yr"] = mmi_val
        out["MMI_Description"] = mmi_label(mmi_val)
    else:
        out["mmi100yr"] = out["MMI_Description"] = "Error" if status != 200 else "N/A"

    # --- US Flood Depth ---
    payload, status, error = li_api.hazard_lookup(
        lat_f, lon_f, admin1_code, "US", resolution_code, "us_fl_depth", body_style="flood_basic",
        layer_options={"lookup": "floodDepth"})
    if status == 200 and isinstance(payload, dict):
        out["floodDepthUndef100yr"] = payload.get("depthInlandUndef100yr", "N/A")
        out["floodDepthDef100yr"] = payload.get("depthInlandDef100yr", "N/A")
        out["floodDepthUndef500yr"] = payload.get("depthInlandUndef500yr", "N/A")
        out["floodDepthDef500yr"] = payload.get("depthInlandDef500yr", "N/A")
    else:
        for k in ("floodDepthUndef100yr", "floodDepthDef100yr", "floodDepthUndef500yr", "floodDepthDef500yr"):
            out[k] = "Error"

    # --- FEMA FIRM flood zone ---
    payload, status, error = li_api.hazard_lookup(
        lat_f, lon_f, admin1_code, "US", resolution_code, "us_fema_firm", body_style="flood_basic",
        layer_options={"distanceUnit": "feet", "proximityDistance": 100})
    if status == 200 and isinstance(payload, dict):
        zone = payload.get("floodZone")
        out["floodZone"] = zone if zone else "Unmapped"
    else:
        out["floodZone"] = "Error"

    # --- US Flood Risk Score (needs building characteristics) ---
    characteristics = {
        "construction": f"ATC{row.get('Construction Code') or 0}",
        "occupancy": f"ATC{row.get('Occupancy Code') or 0}",
        "yearBuilt": row.get("Year Built") if isinstance(row.get("Year Built"), (int, float)) else 0,
        "numOfStories": row.get("# of Stories") if isinstance(row.get("# of Stories"), (int, float)) else 0,
    }
    coverage_values = {
        "buildingValue": row.get("Buildings TIV") or 0,
        "contentsValue": row.get("Contents TIV") or 0,
        "businessInterruptionValue": row.get("BI TIV") or 0,
    }
    payload, status, error = li_api.hazard_lookup(
        lat_f, lon_f, admin1_code, "US", resolution_code, "us_fl_risk_score", body_style="full",
        characteristics=characteristics, coverage_values=coverage_values)
    if status == 200 and isinstance(payload, dict):
        out["floodScoreOverall"] = payload.get("scoreOverall", "N/A")
        out["floodScore100yr"] = payload.get("score100yr", "N/A")
        out["floodScore250yr"] = payload.get("score250yr", "N/A")
        out["floodScore500yr"] = payload.get("score500yr", "N/A")
    else:
        for k in ("floodScoreOverall", "floodScore100yr", "floodScore250yr", "floodScore500yr"):
            out[k] = "Error"

    # --- Geocode (for CRESTA / zone1Code) ---
    payload, status, error = li_api.geocode_flat(
        row.get("Street Address"), row.get("City"), admin1_code, row.get("Postal Code"), "US")
    if status == 200 and isinstance(payload, dict):
        cresta = payload.get("zone1Code")
        out["CRESTA"] = cresta if cresta else "N/A"
    else:
        out["CRESTA"] = "Error"

    return out


def enrich_all_rows(li_api, rows, fallback_df, zip_df, account_number, max_workers, progress_cb=None):
    results = [None] * len(rows)

    def _enrich_one(i, row):
        return i, enrich_row_with_hazards(li_api, row, fallback_df, zip_df, account_number)

    completed = 0
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [executor.submit(_enrich_one, i, row) for i, row in enumerate(rows)]
        for future in as_completed(futures):
            i, hazard_fields = future.result()
            results[i] = hazard_fields
            completed += 1
            if progress_cb:
                progress_cb(completed, len(rows))

    merged = []
    for row, hazard_fields in zip(rows, results):
        new_row = dict(row)
        new_row.update(hazard_fields)
        merged.append(new_row)
    return merged


# --------------------------------------------------------------------------
# Map / summary aggregation (ported from LH_WriteMapSheet)
# --------------------------------------------------------------------------

def build_map_summary(rows_df):
    """Groups by unique Lat,Long - sums TIV, counts distinct locations.
    Mirrors mLH_MapStats in the VBA (keyed by 'lat,lon' string)."""
    if rows_df.empty:
        return pd.DataFrame(columns=["Latitude", "Longitude", "Location TIV Sum", "LocID Count"])
    valid = rows_df[(rows_df["Latitude"].astype(str) != "0") & (rows_df["Longitude"].astype(str) != "0")]
    if valid.empty:
        return pd.DataFrame(columns=["Latitude", "Longitude", "Location TIV Sum", "LocID Count"])
    grouped = valid.groupby(["Latitude", "Longitude"], as_index=False).agg(
        **{"Location TIV Sum": ("Location TIV", "sum"), "LocID Count": ("LocID", "nunique")}
    )
    return grouped


def to_excel_bytes(locations_df, map_df):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        locations_df.to_excel(writer, sheet_name="Locations", index=False)
        map_df.to_excel(writer, sheet_name="Map", index=False)
    buffer.seek(0)
    return buffer.getvalue()


# --------------------------------------------------------------------------
# Streamlit UI
# --------------------------------------------------------------------------

st.set_page_config(page_title="RMS Locations + Hazard Explorer", page_icon="🏢", layout="wide")

st.markdown(
    """
    <style>
    .block-container { padding-top: 2rem; max-width: 1300px; }
    div[data-testid="stExpander"] { border: 1px solid rgba(135,135,135,0.25); border-radius: 10px; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("🏢 RMS Locations + Hazard Enrichment Explorer")
st.caption(
    "Pulls every location for an Exposure ID's latest submission, then enriches each US location "
    "with earthquake, flood, and CRESTA hazard data. Ported from a confirmed-working RMS VBA macro."
)

if "lh_rows" not in st.session_state:
    st.session_state["lh_rows"] = None
if "lh_stage" not in st.session_state:
    st.session_state["lh_stage"] = None

with st.sidebar:
    st.header("Connection")
    host = st.text_input("Host", value="api-euw1.rms.com")
    riskdata_key = st.text_input("Risk Data API key", type="password",
                                  help="Used for GET .../exposures/{id}/locations and location detail calls.")
    hazard_key = st.text_input("Hazard / Location Intelligence API key", type="password",
                                help="Used for all li/* hazard and geocode calls. This is a SEPARATE key "
                                     "from the Risk Data key, even though both hit the same host.")
    verify_ssl = st.checkbox("Verify SSL", value=True)
    workers = st.slider("Parallel requests (workers)", 1, 20, 8,
                         help="Applies to both the location-detail pull and the hazard enrichment stage.")

    st.divider()
    st.subheader("Optional: Admin1 (state) fallback")
    st.caption(
        "If a location's address is missing a resolvable state, these optional lookups fill the gap - "
        "equivalent to the macro's LOC-sheet and uszips-sheet fallback chain."
    )
    fallback_csv = st.file_uploader(
        "Location state fallback CSV", type=["csv"],
        help="Columns: AccountNumber, LocationNumber, StateCode, PostalCode")
    zip_csv = st.file_uploader(
        "ZIP-to-state CSV", type=["csv"],
        help="Columns: zip, state_id (e.g. the public SimpleMaps uszips.csv)")

fallback_df = None
if fallback_csv is not None:
    try:
        fallback_df = pd.read_csv(fallback_csv, dtype=str)
        st.sidebar.success(f"Loaded {len(fallback_df)} fallback row(s).")
    except Exception as exc:
        st.sidebar.error(f"Could not read fallback CSV: {exc}")

zip_df = None
if zip_csv is not None:
    try:
        zip_df = pd.read_csv(zip_csv, dtype=str)
        zip_df["zip"] = zip_df["zip"].apply(normalize_zip)
        st.sidebar.success(f"Loaded {len(zip_df)} ZIP row(s).")
    except Exception as exc:
        st.sidebar.error(f"Could not read ZIP CSV: {exc}")

st.subheader("1. Exposure")
col1, col2 = st.columns([2, 2])
with col1:
    exposure_id = st.text_input("Exposure ID", placeholder="2918633")
with col2:
    account_number = st.text_input("Account Name/Number (label only)", placeholder="ACC-205")

run = st.button("Pull locations & enrich hazards", type="primary")

if run:
    if not riskdata_key or not hazard_key:
        st.warning("Please provide both API keys (sidebar) before running.")
    elif not exposure_id.strip():
        st.warning("Please enter an Exposure ID.")
    else:
        rd_api = RiskDataApi(host.strip(), riskdata_key, verify=verify_ssl)
        li_api = LocationIntelligenceApi(host.strip(), hazard_key, verify=verify_ssl)

        progress = st.progress(0.0, text="Pulling locations...")
        rows, error = pull_locations_for_exposure(
            rd_api, exposure_id.strip(), account_number or exposure_id.strip(), workers,
            progress_cb=lambda done, total: progress.progress(
                min(done / max(total, 1), 1.0) * 0.5, text=f"Stage 1/2 - fetching location detail: {done}/{total}"),
        )

        if error:
            progress.empty()
            st.error(error)
            st.session_state["lh_rows"] = None
        else:
            progress.progress(0.5, text=f"Stage 2/2 - enriching {len(rows)} location(s) with hazard data...")
            enriched = enrich_all_rows(
                li_api, rows, fallback_df, zip_df, account_number or exposure_id.strip(), workers,
                progress_cb=lambda done, total: progress.progress(
                    0.5 + min(done / max(total, 1), 1.0) * 0.5,
                    text=f"Stage 2/2 - enriching location {done}/{total}..."),
            )
            progress.empty()
            st.session_state["lh_rows"] = enriched
            st.success(f"Done - {len(enriched)} location(s) pulled and enriched.")

# --------------------------------------------------------------------------
# Results
# --------------------------------------------------------------------------

rows = st.session_state["lh_rows"]
if rows:
    df = pd.DataFrame(rows)
    display_cols = [c for c in LOCATION_CORE_HEADERS + HAZARD_HEADERS if c in df.columns]
    df_display = df[display_cols].copy()

    st.divider()
    st.subheader("📊 Summary")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Locations", len(df_display))
    m2.metric("Total TIV", f"${df_display['Location TIV'].sum():,.0f}")
    hazard_ok = (df_display.get("floodZone", pd.Series(dtype=str)) != "Error").sum() if "floodZone" in df_display else 0
    m3.metric("Hazard enrichment OK", f"{hazard_ok}/{len(df_display)}")
    cresta_zones = df_display["CRESTA"].replace({"": None, "N/A": None, "Error": None}).dropna().nunique() if "CRESTA" in df_display else 0
    m4.metric("Distinct CRESTA zones", cresta_zones)

    st.divider()
    st.subheader("🗺️ Map")
    map_source = df_display[["Latitude", "Longitude"]].copy()
    map_source["Latitude"] = pd.to_numeric(map_source["Latitude"], errors="coerce")
    map_source["Longitude"] = pd.to_numeric(map_source["Longitude"], errors="coerce")
    map_source = map_source.dropna()
    map_source = map_source.rename(columns={"Latitude": "lat", "Longitude": "lon"})
    if not map_source.empty:
        st.map(map_source, zoom=4)
    else:
        st.info("No valid coordinates to plot.")

    st.divider()
    st.subheader("📋 Locations")
    st.dataframe(df_display, width="stretch", hide_index=True)

    map_df = build_map_summary(df_display.assign(
        Latitude=pd.to_numeric(df_display["Latitude"], errors="coerce").fillna(0),
        Longitude=pd.to_numeric(df_display["Longitude"], errors="coerce").fillna(0),
    ))

    st.divider()
    st.subheader("⬇️ Export")
    c1, c2 = st.columns(2)
    with c1:
        st.download_button("Download Locations as CSV", data=df_display.to_csv(index=False).encode("utf-8"),
                            file_name="rms_locations_hazard.csv", mime="text/csv", width="stretch")
    with c2:
        st.download_button("Download Excel (Locations + Map)",
                            data=to_excel_bytes(df_display, map_df),
                            file_name="rms_locations_hazard.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            width="stretch")
