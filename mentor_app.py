import streamlit as st
import pandas as pd
import plotly.express as px
import requests
import zipfile
import io
import os
import base64
import subprocess
from bs4 import BeautifulSoup

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

INCIDENT_COLUMNS = [
    "year", "state", "operator", "cause", "cause_details", "narrative",
    "fatalities", "damage_cost", "incident_date", "city", "county", "location",
]

APP_NAME = "SKIM"
APP_FULL_NAME = "Southern Knowledge Internal Mentor"
APP_TAGLINE = "Southern Company · Internal Knowledge Assistant"
_APP_DIR = os.path.dirname(os.path.abspath(__file__))
LOGO_SVG_PATH = os.path.join(_APP_DIR, "assets", "southern_company_logo.svg")
LOGO_PNG_PATH = os.path.join(_APP_DIR, "assets", "southern_company_logo.png")

TERRITORY_STATES = ["GA", "AL", "TN", "MS", "FL", "VA", "NC"]

OPERATOR_KEYWORDS = ["southern", "atlanta gas", "chattanooga gas"]

COLUMN_MAP = {
    "iyear": "year",
    "year": "year",
    "report_year": "year",
    "state_code": "state",
    "state": "state",
    "operator_name": "operator",
    "operator": "operator",
    "cause": "cause",
    "cause_of_failure": "cause",
    "cause_details": "cause_details",
    "significant_cause": "cause_details",
    "narrative": "narrative",
    "accident_description": "narrative",
    "accident_summary": "narrative",
    "injuries_and_fatalities": "fatalities",
    "tot_fatal": "fatalities",
    "fatalities_total": "fatalities",
    "num_fatalities": "fatalities",
    "total_cost_in_dollars": "damage_cost",
    "total_cost": "damage_cost",
    "property_damage_cost": "damage_cost",
    "total_cost_current": "damage_cost",
    "location_state_abbreviation": "state",
    "location_city_name": "city",
    "location_county_name": "county",
    "name": "operator",
    "fatal": "fatalities",
    "incident_date_time": "incident_date",
    "local_date": "incident_date",
    "incident_date": "incident_date",
    "city_name": "city",
    "city": "city",
    "county_name": "county",
    "county": "county",
    "pipeline_location": "location",
    "location": "location",
}

PHMSA_INCIDENT_PAGE = (
    "https://www.phmsa.dot.gov/data-and-statistics/pipeline/"
    "distribution-transmission-gathering-lng-and-liquid-accident-and-incident-data"
)
# RETIRED: DOT removed jzjf-e6ij from Socrata (returns dataset.missing / Not found).
# PHMSA gas-distribution ZIP on phmsa.dot.gov is the current live federal source.
SOCRATA_LEGACY_DATASET_ID = "jzjf-e6ij"
RETIRED_DATASET_MSG = (
    "Socrata dataset jzjf-e6ij was retired by DOT (dataset.missing). "
    "Live incident data now comes from the PHMSA gas-distribution ZIP on phmsa.dot.gov."
)
SOCRATA_FALLBACK_DATASET_ID = "27nc-rsge"
# Official DOT Pipeline Incident Flagged Files (gas distribution + other systems)
DOT_FLAGGED_FILES_ID = "qdme-9bbm"
# Live DOT Socrata datasets to try (bulk download, filter locally)
DOT_INCIDENT_DATASET_IDS = [
    ("qdme-9bbm", "DOT Pipeline Incident Flagged Files"),
    ("27nc-rsge", "DOT Combined Pipeline Incidents"),
]
SOCRATA_CATALOG_URL = (
    "https://api.us.socrata.com/api/catalog/v1"
    "?domains=data.transportation.gov"
    "&q=pipeline+gas+distribution+incident"
    "&only=datasets&limit=15"
)
SOCRATA_APP_TOKEN = os.getenv("SOCRATA_APP_TOKEN", "").strip()

# User-downloaded PHMSA files (browser download when .gov blocks scripts)
LOCAL_PHMSA_CANDIDATES = [
    os.getenv("PHMSA_LOCAL_ZIP", "").strip(),
    os.getenv("PHMSA_LOCAL_DIR", "").strip(),
    os.path.expanduser("~/Downloads/PHMSA_Pipeline_Safety_Flagged_Incidents"),
    os.path.expanduser("~/Downloads/PHMSA_Pipeline_Safety_Flagged_Incidents.zip"),
    os.path.expanduser("~/Downloads/incident_gas_distribution_jan2024.zip"),
    os.path.expanduser("~/Downloads/incident_gas_distribution_jan2023.zip"),
]
# Gas distribution data sheets inside the flagged-files ZIP
FLAGGED_GD_SHEETS = [
    ("gd2010toPresent.xlsx", "gd2010toPresent"),
    ("gdmar2004to2009.xlsx", "gdmar2004to2009"),
    ("gd1986tofeb2004.xlsx", "gd1986tofeb2004"),
]

HTTP_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/122.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json,text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Accept-Encoding": "gzip, deflate",
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1",
}

SEC_HEADERS = {
    "User-Agent": "SKIM southern-company-gas@example.com",
    "Accept": "application/json",
    "Accept-Encoding": "gzip, deflate",
}

_HTTP_SESSION = None


def _get_http_session():
    global _HTTP_SESSION
    if _HTTP_SESSION is None:
        _HTTP_SESSION = requests.Session()
        _HTTP_SESSION.trust_env = False
        _HTTP_SESSION.headers.update(HTTP_HEADERS)
    return _HTTP_SESSION


def _request_headers(extra=None):
    headers = dict(HTTP_HEADERS)
    if SOCRATA_APP_TOKEN:
        headers["X-App-Token"] = SOCRATA_APP_TOKEN
    if extra:
        headers.update(extra)
    return headers


def _http_get(url, **kwargs):
    kwargs.setdefault("timeout", 20)
    kwargs.setdefault("proxies", {"http": None, "https": None})
    extra = kwargs.pop("headers", None)
    kwargs["headers"] = _request_headers(extra)
    return _get_http_session().get(url, **kwargs)


def _warm_session(url, referer=None):
    """Visit a landing page so subsequent downloads may inherit session cookies."""
    try:
        _http_get(url, timeout=25, headers={"Referer": referer or url})
    except Exception:
        pass


def _curl_download(url, referer=None, timeout=120):
    """Fallback transport when requests gets HTTP 403 from .gov WAF."""
    cmd = [
        "curl", "-fsSL", "--max-time", str(timeout),
        "-A", HTTP_HEADERS["User-Agent"],
        "-H", "Accept: */*",
        url,
    ]
    if referer:
        cmd.extend(["-e", referer])
    if SOCRATA_APP_TOKEN:
        cmd.extend(["-H", f"X-App-Token: {SOCRATA_APP_TOKEN}"])
    result = subprocess.run(cmd, capture_output=True, check=False)
    if result.returncode != 0:
        err = (result.stderr or result.stdout or b"").decode("utf-8", errors="replace")[:200]
        raise RuntimeError(err or f"curl failed (exit {result.returncode})")
    return result.stdout


def _read_csv_bytes(content):
    for enc in ("utf-8", "latin-1", "cp1252"):
        try:
            return pd.read_csv(io.BytesIO(content), low_memory=False, encoding=enc)
        except UnicodeDecodeError:
            continue
    return pd.read_csv(io.BytesIO(content), low_memory=False)


def _absolute_url(href, base="https://www.phmsa.dot.gov"):
    href = (href or "").strip()
    if not href:
        return ""
    if href.startswith("/"):
        return base.rstrip("/") + href
    if not href.startswith("http"):
        return base.rstrip("/") + "/" + href.lstrip("/")
    return href


def _parse_socrata_payload(data):
    """Normalize Socrata resource JSON or bulk rows.json into a DataFrame."""
    if isinstance(data, list):
        return pd.DataFrame(data)
    if isinstance(data, dict):
        if data.get("error"):
            raise ValueError(data.get("message", str(data)))
        if "data" in data and "meta" in data:
            columns = data.get("meta", {}).get("view", {}).get("columns", [])
            names = [c.get("name", c.get("fieldName", f"col_{i}")) for i, c in enumerate(columns)]
            rows = data.get("data", [])
            if names and rows:
                return pd.DataFrame(rows, columns=names)
        if "results" in data:
            return pd.DataFrame(data["results"])
    return pd.DataFrame()


def _territory_filter_df(raw_df):
    if raw_df is None or raw_df.empty:
        return raw_df
    state_col = next(
        (c for c in raw_df.columns if c.strip().lower() in ("state_code", "state")), None
    )
    if not state_col:
        return raw_df
    mask = raw_df[state_col].astype(str).str.upper().isin(TERRITORY_STATES)
    return raw_df[mask].copy()


def _filter_gas_distribution(raw_df):
    """Keep gas distribution rows when the combined PHMSA dataset includes a type column."""
    if raw_df is None or raw_df.empty:
        return raw_df
    for col in raw_df.columns:
        key = col.strip().lower()
        if key in ("system_type", "type_of_system", "part_of_system", "commodity_group"):
            vals = raw_df[col].astype(str).str.lower()
            mask = vals.str.contains("distrib", na=False)
            if mask.any():
                return raw_df[mask].copy()
    return raw_df


def _discover_phmsa_distribution_zip():
    """Scrape PHMSA incident page for the current gas distribution ZIP download URL."""
    try:
        resp = _http_get(PHMSA_INCIDENT_PAGE, timeout=45, headers={"Referer": "https://www.phmsa.dot.gov/"})
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
        candidates = []
        for a in soup.find_all("a", href=True):
            href = a["href"].strip()
            text = (a.get_text() or "").strip()
            text_lower = text.lower()
            href_lower = href.lower()
            if "zip" not in href_lower:
                continue
            score = 0
            if "incident_gas_distribution" in href_lower:
                score += 5
            if "gas distribution" in text_lower and "incident" in text_lower:
                score += 4
            if "2010" in text_lower or "present" in text_lower:
                score += 2
            if score > 0:
                candidates.append((score, _absolute_url(href), text))
        if candidates:
            candidates.sort(key=lambda x: x[0], reverse=True)
            return candidates[0][1], candidates[0][2]
    except Exception as exc:
        print(f"PHMSA ZIP discovery failed: {exc}")
    return None, None


def _load_flagged_gd_frames(list_names, open_xlsx):
    """Load gas distribution sheets from a ZIP or extracted folder."""
    names = set(list_names())
    frames = []
    for xlsx_name, sheet_name in FLAGGED_GD_SHEETS:
        if xlsx_name not in names:
            continue
        with open_xlsx(xlsx_name) as f:
            frames.append(pd.read_excel(f, sheet_name=sheet_name))
    if not frames:
        raise ValueError("No gas distribution sheets found in flagged PHMSA data")
    return pd.concat(frames, ignore_index=True, sort=False)


def _load_flagged_zip_local(path):
    """Load gas distribution incidents from PHMSA Pipeline Safety Flagged Files ZIP."""
    with zipfile.ZipFile(path) as zf:
        return _load_flagged_gd_frames(zf.namelist, lambda name: zf.open(name))


def _load_flagged_dir_local(path):
    """Load gas distribution incidents from an extracted flagged-files folder."""
    return _load_flagged_gd_frames(
        lambda: os.listdir(path),
        lambda name: open(os.path.join(path, name), "rb"),
    )


def _load_incidents_from_local_path(path):
    """Load incidents from a user-downloaded PHMSA file or folder on disk."""
    if os.path.isdir(path):
        if os.path.isfile(os.path.join(path, "gd2010toPresent.xlsx")):
            return _load_flagged_dir_local(path)
        raise ValueError(f"No gas distribution sheets found in folder: {path}")
    lower = path.lower()
    if lower.endswith(".zip"):
        with zipfile.ZipFile(path) as zf:
            if "gd2010toPresent.xlsx" in zf.namelist():
                return _load_flagged_zip_local(path)
            csv_names = [n for n in zf.namelist() if n.lower().endswith(".csv")]
            if csv_names:
                with zf.open(max(csv_names, key=len)) as f:
                    return pd.read_csv(f, low_memory=False, encoding="latin-1")
            data_names = [
                n for n in zf.namelist()
                if n.lower().endswith((".txt", ".tsv")) and "field" not in n.lower()
            ]
            if data_names:
                with zf.open(data_names[0]) as f:
                    return pd.read_csv(f, sep="\t", low_memory=False, encoding="latin-1")
        raise ValueError("No incident data found in local ZIP")
    if lower.endswith((".xlsx", ".xls")):
        xl = pd.ExcelFile(path)
        sheet = "gd2010toPresent" if "gd2010toPresent" in xl.sheet_names else xl.sheet_names[-1]
        return pd.read_excel(path, sheet_name=sheet)
    if lower.endswith(".csv"):
        return _read_csv_bytes(open(path, "rb").read())
    raise ValueError(f"Unsupported local file type: {path}")


def _fetch_local_phmsa_incidents(errors):
    """Path 0 — user-downloaded PHMSA file (real federal data, bypasses .gov 403)."""
    for path in LOCAL_PHMSA_CANDIDATES:
        if not path or not (os.path.isfile(path) or os.path.isdir(path)):
            continue
        try:
            raw = _load_incidents_from_local_path(path)
            df = _normalize_incidents_df(raw)
            if len(df) > 0:
                label = f"Local PHMSA file ({os.path.basename(path.rstrip(os.sep))})"
                print(f"PHMSA incidents: {len(df)} rows from {label}")
                return df, label
            errors.append(f"{path}: loaded {len(raw)} rows, none matched territory filter")
        except Exception as exc:
            errors.append(f"Local file {path}: {exc}")
    return None, None


def _load_incidents_from_zip(zip_url, use_curl=False):
    """Download a PHMSA ZIP and return the raw incidents DataFrame."""
    referer = PHMSA_INCIDENT_PAGE
    if use_curl:
        content = _curl_download(zip_url, referer=referer, timeout=120)
    else:
        resp = _http_get(zip_url, timeout=90, headers={"Referer": referer})
        if resp.status_code == 403:
            content = _curl_download(zip_url, referer=referer, timeout=120)
        else:
            resp.raise_for_status()
            content = resp.content
    with zipfile.ZipFile(io.BytesIO(content)) as zf:
        csv_names = [n for n in zf.namelist() if n.lower().endswith(".csv")]
        if not csv_names:
            data_names = [
                n for n in zf.namelist()
                if n.lower().endswith((".txt", ".tsv")) and "field" not in n.lower()
            ]
            if not data_names:
                raise ValueError("No data file found in PHMSA ZIP")
            with zf.open(data_names[0]) as f:
                return pd.read_csv(f, sep="\t", low_memory=False, encoding="latin-1")
        data_csv = max(csv_names, key=len)
        with zf.open(data_csv) as f:
            return pd.read_csv(f, low_memory=False, encoding="latin-1")


def _socrata_referer_headers(dataset_id):
    return {
        "Referer": f"https://data.transportation.gov/resource/{dataset_id}",
        "Origin": "https://data.transportation.gov",
    }


def _fetch_socrata_legacy_incidents(errors):
    """Probe retired Endpoint A — jzjf-e6ij (DOT removed this dataset)."""
    dataset_id = SOCRATA_LEGACY_DATASET_ID
    url = f"https://data.transportation.gov/resource/{dataset_id}.json"
    params = {"$limit": 1}
    label = f"Socrata API ({dataset_id})"
    try:
        resp = _http_get(
            url, params=params, timeout=15,
            headers=_socrata_referer_headers(dataset_id),
        )
        if resp.status_code == 404:
            errors.append(f"{label}: retired (HTTP 404)")
            errors.insert(0, RETIRED_DATASET_MSG)
            return None, None
        if resp.status_code == 403:
            errors.append(f"{label}: HTTP 403 Forbidden")
            return None, None
        resp.raise_for_status()
        data = resp.json()
        if isinstance(data, dict) and data.get("code") == "dataset.missing":
            errors.insert(0, RETIRED_DATASET_MSG)
            errors.append(f"{label}: {data.get('message', 'dataset.missing')}")
            return None, None
        if isinstance(data, dict) and data.get("error"):
            errors.append(f"{label}: {data.get('message', data)}")
            return None, None
        if not data:
            errors.insert(0, RETIRED_DATASET_MSG)
            return None, None
        # Unexpected: dataset came back — use full territory query
        resp2 = _http_get(
            url,
            params={
                "$limit": 2000,
                "$where": "state_code in ('GA','AL','TN','MS','FL','VA','NC')",
                "$order": "iyear DESC",
            },
            timeout=30,
            headers=_socrata_referer_headers(dataset_id),
        )
        resp2.raise_for_status()
        df = _normalize_incidents_df(pd.DataFrame(resp2.json()))
        if len(df) > 0:
            return df, label
    except Exception as exc:
        errors.append(f"{label}: {exc}")
    return None, None


def _incidents_from_raw_frame(raw, source_label):
    """Filter and normalize a raw incident DataFrame."""
    if raw is None or raw.empty:
        return None
    raw = _territory_filter_df(raw)
    raw = _filter_gas_distribution(raw)
    df = _normalize_incidents_df(raw)
    if len(df) > 0:
        return df, source_label
    return None


def _fetch_dot_bulk_dataset(dataset_id, label, errors, use_curl=False):
    """Download a full DOT/Socrata dataset export and filter territory locally."""
    referer = _socrata_referer_headers(dataset_id)
    page = f"https://data.transportation.gov/d/{dataset_id}"
    urls = [
        (f"https://data.transportation.gov/api/views/{dataset_id}/rows.csv", {"accessType": "DOWNLOAD"}),
        (f"https://data.transportation.gov/api/views/{dataset_id}/rows.csv", {}),
        (f"https://data.transportation.gov/resource/{dataset_id}.csv", {"$limit": 50000}),
        (f"https://data.transportation.gov/resource/{dataset_id}.csv", {}),
    ]
    for idx, (url, params) in enumerate(urls, start=1):
        attempt = f"{label} export {idx}"
        try:
            if use_curl:
                qs = "&".join(f"{k}={v}" for k, v in params.items()) if params else ""
                full_url = f"{url}?{qs}" if qs else url
                raw_bytes = _curl_download(full_url, referer=page, timeout=180)
                raw = _read_csv_bytes(raw_bytes)
            else:
                resp = _http_get(url, params=params or None, timeout=180, headers=referer)
                if resp.status_code == 403 and not use_curl:
                    raw_bytes = _curl_download(
                        resp.url, referer=page, timeout=180
                    )
                    raw = _read_csv_bytes(raw_bytes)
                else:
                    if resp.status_code in (400, 403, 404):
                        errors.append(f"{attempt}: HTTP {resp.status_code}")
                        continue
                    resp.raise_for_status()
                    if "json" in url:
                        raw = _parse_socrata_payload(resp.json())
                    else:
                        raw = _read_csv_bytes(resp.content)
            result = _incidents_from_raw_frame(raw, f"{attempt} ({dataset_id})")
            if result:
                return result
            errors.append(f"{attempt}: {len(raw)} rows downloaded, none in territory")
        except Exception as exc:
            errors.append(f"{attempt}: {exc}")
    return None, None


def _discover_dot_incident_datasets(errors):
    """Query Socrata catalog for current pipeline incident dataset IDs."""
    found = []
    try:
        resp = _http_get(SOCRATA_CATALOG_URL, timeout=30)
        if resp.status_code != 200:
            errors.append(f"Socrata catalog: HTTP {resp.status_code}")
            return found
        payload = resp.json()
        for item in payload.get("results", []):
            resource = item.get("resource", {})
            ds_id = resource.get("id", "")
            name = resource.get("name", "")
            if not ds_id or len(ds_id) != 9:
                continue
            if ds_id in (SOCRATA_LEGACY_DATASET_ID,):
                continue
            if any(x in name.lower() for x in ("incident", "accident", "pipeline", "phmsa")):
                found.append((ds_id, name[:80]))
    except Exception as exc:
        errors.append(f"Socrata catalog discovery: {exc}")
    return found


def _fetch_all_dot_datasets(errors):
    """Try every known + discovered DOT incident dataset via bulk export."""
    seen = set()
    candidates = list(DOT_INCIDENT_DATASET_IDS)
    for ds_id, name in _discover_dot_incident_datasets(errors):
        if ds_id not in seen:
            candidates.append((ds_id, name))

    for dataset_id, label in candidates:
        if dataset_id in seen:
            continue
        seen.add(dataset_id)
        for use_curl in (False, True):
            df, source = _fetch_dot_bulk_dataset(dataset_id, label, errors, use_curl=use_curl)
            if df is not None and len(df) > 0:
                return df, source
    return None, None


def _empty_incidents_df():
    return pd.DataFrame(columns=INCIDENT_COLUMNS)


def _coalesce_duplicate_columns(df):
    """Merge columns that share the same name after PHMSA field mapping."""
    if not df.columns.duplicated().any():
        return df
    out = {}
    for col in dict.fromkeys(df.columns):
        parts = [df.iloc[:, i] for i, c in enumerate(df.columns) if c == col]
        if len(parts) == 1:
            out[col] = parts[0]
        else:
            combined = parts[0]
            for part in parts[1:]:
                combined = combined.where(
                    combined.notna() & (combined.astype(str).str.strip() != ""),
                    part,
                )
            out[col] = combined
    return pd.DataFrame(out)


def _normalize_incidents_df(raw_df):
    """Rename columns, coerce types, fill NaN, and filter to territory."""
    if raw_df is None or raw_df.empty:
        return _empty_incidents_df()

    try:
        df = raw_df.copy()
        rename = {}
        for col in df.columns:
            key = col.strip().lower()
            if key in COLUMN_MAP:
                rename[col] = COLUMN_MAP[key]
        df = df.rename(columns=rename)
        df = _coalesce_duplicate_columns(df)

        for col in INCIDENT_COLUMNS:
            if col not in df.columns:
                df[col] = ""

        if "year" in df.columns:
            df["year"] = pd.to_numeric(df["year"], errors="coerce").fillna(0).astype(int)
        if "damage_cost" in df.columns:
            df["damage_cost"] = pd.to_numeric(df["damage_cost"], errors="coerce").fillna(0.0)
        if "fatalities" in df.columns:
            df["fatalities"] = pd.to_numeric(df["fatalities"], errors="coerce").fillna(0.0)

        for col in INCIDENT_COLUMNS:
            if col in ("year", "damage_cost", "fatalities"):
                continue
            df[col] = df[col].fillna("").astype(str)

        op_lower = df["operator"].str.lower()
        state_upper = df["state"].str.upper()
        operator_match = False
        for kw in OPERATOR_KEYWORDS:
            operator_match = operator_match | op_lower.str.contains(kw, na=False)
        territory_match = state_upper.isin(TERRITORY_STATES)
        df = df[operator_match | territory_match].reset_index(drop=True)
        return df[INCIDENT_COLUMNS]
    except Exception:
        return _empty_incidents_df()


# ---------------------------------------------------------------------------
# PART 1 — Data fetching
# ---------------------------------------------------------------------------

@st.cache_data(ttl=3600, show_spinner=False)
def fetch_phmsa_incidents():
    """Fetch real PHMSA gas distribution incidents for Southern Company territory."""
    errors = []
    _warm_session("https://www.phmsa.dot.gov/", referer="https://www.phmsa.dot.gov/")
    _warm_session("https://data.transportation.gov/", referer="https://data.transportation.gov/")

    # Path 0 — user-downloaded PHMSA ZIP in ~/Downloads (bypasses .gov 403)
    df, source = _fetch_local_phmsa_incidents(errors)
    if df is not None and len(df) > 0:
        return df, errors

    # Path 1 — PHMSA gas-distribution ZIP via network
    zip_url, zip_label = _discover_phmsa_distribution_zip()
    zip_candidates = []
    if zip_url:
        zip_candidates.append((zip_url, zip_label or "PHMSA discovered ZIP"))
    zip_candidates.extend([
        (
            "https://www.phmsa.dot.gov/sites/phmsa.dot.gov/files/"
            "2024-10/incident_gas_distribution_jan2024.zip",
            "PHMSA gas distribution 2010–present",
        ),
        (
            "https://www.phmsa.dot.gov/sites/phmsa.dot.gov/files/"
            "2023-10/incident_gas_distribution_jan2023.zip",
            "PHMSA gas distribution fallback 2023",
        ),
    ])
    for url, label in zip_candidates:
        for use_curl in (False, True):
            try:
                raw = _load_incidents_from_zip(url, use_curl=use_curl)
                df = _normalize_incidents_df(raw)
                if len(df) > 0:
                    via = "curl" if use_curl else "https"
                    print(f"PHMSA incidents: {len(df)} rows from {label} ({via})")
                    return df, errors
            except Exception as exc:
                errors.append(f"{label}: {exc}")

    # Path 2 — DOT Open Data bulk exports (qdme-9bbm flagged files, 27nc-rsge, catalog)
    df, source = _fetch_all_dot_datasets(errors)
    if df is not None and len(df) > 0:
        print(f"PHMSA incidents: loaded {len(df)} rows from {source}")
        return df, errors

    # Path 3 — confirm legacy dataset is retired
    _fetch_socrata_legacy_incidents(errors)

    if not errors:
        errors.append("All PHMSA incident sources returned zero rows")
    print("PHMSA incidents: total failure — returning empty DataFrame")
    return _empty_incidents_df(), errors


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_phmsa_annual_stats():
    """Fetch PHMSA annual operator statistics for Southern Company Gas."""
    empty = pd.DataFrame(columns=["year", "operator", "pipeline_miles", "services", "surveyed"])

    # Endpoint A — Socrata API (legacy ID may be retired; try then fall through)
    try:
        url = "https://data.transportation.gov/resource/myei-c3fa.json"
        params = {
            "$limit": 500,
            "$where": (
                "operator_name like '%Southern%' OR "
                "operator_name like '%Atlanta Gas%'"
            ),
            "$order": "report_year DESC",
        }
        resp = _http_get(url, params=params, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        if data and not (isinstance(data, dict) and data.get("error")):
            df = pd.DataFrame(data)
            rename = {
                "report_year": "year",
                "total_miles": "pipeline_miles",
                "services_inst": "services",
                "leak_survey_miles": "surveyed",
            }
            df = df.rename(columns={k: v for k, v in rename.items() if k in df.columns})
            if "operator_name" in df.columns and "operator" not in df.columns:
                df = df.rename(columns={"operator_name": "operator"})
            for col in ["year", "operator", "pipeline_miles", "services", "surveyed"]:
                if col not in df.columns:
                    df[col] = ""
            if "year" in df.columns:
                df["year"] = pd.to_numeric(df["year"], errors="coerce").fillna(0).astype(int)
            for col in ("pipeline_miles", "services", "surveyed"):
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
            if len(df) > 0:
                print(f"PHMSA annual stats: loaded {len(df)} rows from Socrata API")
                return df[["year", "operator", "pipeline_miles", "services", "surveyed"]]
    except Exception as exc:
        print(f"PHMSA annual stats Socrata failed: {exc}")

    # Endpoint B — PHMSA direct ZIP download
    try:
        zip_url = (
            "https://www.phmsa.dot.gov/sites/phmsa.dot.gov/files/"
            "2024-10/annual_gas_distribution_jan2024.zip"
        )
        resp = _http_get(zip_url, timeout=60)
        resp.raise_for_status()
        with zipfile.ZipFile(io.BytesIO(resp.content)) as zf:
            csv_names = [n for n in zf.namelist() if n.lower().endswith(".csv")]
            if not csv_names:
                raise ValueError("No CSV in annual ZIP")
            with zf.open(csv_names[0]) as f:
                raw = pd.read_csv(f, low_memory=False)

        op_col = next(
            (c for c in raw.columns if c.lower() in ("operator_name", "operator")), None
        )
        if op_col:
            mask = raw[op_col].astype(str).str.lower().str.contains(
                "southern|atlanta gas", na=False, regex=True
            )
            raw = raw[mask]

        rename = {
            "report_year": "year",
            "operator_name": "operator",
            "total_miles": "pipeline_miles",
            "services_inst": "services",
            "leak_survey_miles": "surveyed",
        }
        df = raw.rename(columns={k: v for k, v in rename.items() if k in raw.columns})
        for col in ["year", "operator", "pipeline_miles", "services", "surveyed"]:
            if col not in df.columns:
                df[col] = ""
        if "year" in df.columns:
            df["year"] = pd.to_numeric(df["year"], errors="coerce").fillna(0).astype(int)
        for col in ("pipeline_miles", "services", "surveyed"):
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
        if len(df) > 0:
            print(f"PHMSA annual stats: loaded {len(df)} rows from PHMSA ZIP")
            return df[["year", "operator", "pipeline_miles", "services", "surveyed"]]
    except Exception as exc:
        print(f"PHMSA annual stats ZIP failed: {exc}")

    return empty


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_sec_filings():
    """Fetch Southern Company SEC filings index from EDGAR."""
    empty = pd.DataFrame(columns=["filing_date", "form", "description", "accession"])
    try:
        url = "https://data.sec.gov/submissions/CIK0000092122.json"
        resp = _http_get(url, headers=SEC_HEADERS, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        recent = data.get("filings", {}).get("recent", {})
        if not recent:
            return empty

        df = pd.DataFrame(recent)
        keep = [c for c in ["filingDate", "form", "primaryDocDescription", "accessionNumber"]
                if c in df.columns]
        df = df[keep].head(50)
        df = df.rename(columns={
            "filingDate": "filing_date",
            "form": "form",
            "primaryDocDescription": "description",
            "accessionNumber": "accession",
        })
        print(f"SEC filings: loaded {len(df)} rows")
        return df
    except Exception as exc:
        print(f"SEC filings failed: {exc}")
        return empty


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_ferc_gas_data():
    """Scrape FERC Form 2 historical data page and load first available CSV."""
    empty = pd.DataFrame()
    try:
        page_url = (
            "https://www.ferc.gov/industries-data/natural-gas/"
            "industry-forms/form-2-2a-3-q-gas-historical-vfp-data"
        )
        resp = _http_get(page_url, timeout=30, headers={"Referer": "https://www.ferc.gov/"})
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")

        csv_links = []
        for a in soup.find_all("a", href=True):
            href = _absolute_url(a["href"], base="https://www.ferc.gov")
            if href and "csv" in href.lower():
                csv_links.append(href)

        zip_links = []
        for a in soup.find_all("a", href=True):
            href = _absolute_url(a["href"], base="https://www.ferc.gov")
            if href and "zip" in href.lower():
                zip_links.append(href)

        if not csv_links and zip_links:
            zresp = _http_get(zip_links[0], timeout=60)
            zresp.raise_for_status()
            with zipfile.ZipFile(io.BytesIO(zresp.content)) as zf:
                csv_in_zip = [n for n in zf.namelist() if n.lower().endswith(".csv")]
                if csv_in_zip:
                    with zf.open(csv_in_zip[0]) as f:
                        df = pd.read_csv(f, low_memory=False, nrows=200)
                        print(f"FERC data: loaded {len(df)} rows from ZIP")
                        return df.head(200)

        for link in csv_links:
            try:
                cresp = _http_get(link, timeout=60)
                cresp.raise_for_status()
                df = pd.read_csv(io.BytesIO(cresp.content), low_memory=False, nrows=200)
                print(f"FERC data: loaded {len(df)} rows from {link}")
                return df.head(200)
            except Exception:
                continue

        print("FERC data: no downloadable CSV found")
        return empty
    except Exception as exc:
        print(f"FERC data failed: {exc}")
        return empty


# ---------------------------------------------------------------------------
# Resolution store helpers
# ---------------------------------------------------------------------------

def extract_symptoms(narrative, cause):
    symptoms = []
    narrative_lower = str(narrative).lower()
    cause_lower = str(cause).lower()
    if any(w in narrative_lower for w in ["pressure", "psi", "drop"]):
        symptoms.append("Pressure anomaly detected")
    if any(w in narrative_lower for w in ["smell", "odor", "gas odor"]):
        symptoms.append("Gas odor reported")
    if any(w in narrative_lower for w in ["leak", "leaking", "escaping"]):
        symptoms.append("Active gas leak")
    if any(w in narrative_lower for w in ["excavat", "dig", "bore"]):
        symptoms.append("Excavation activity nearby")
    if any(w in narrative_lower for w in ["corros", "rust", "deteriorat"]):
        symptoms.append("Corrosion signs")
    if any(w in cause_lower for w in ["material", "joint", "weld"]):
        symptoms.append("Material or joint failure")
    if not symptoms:
        symptoms.append(f"Incident related to {cause}")
    return symptoms


def map_cause_to_regulation(cause):
    cause_lower = str(cause).lower()
    if "corros" in cause_lower:
        return "49 CFR 192.481 — Atmospheric Corrosion"
    if "excavat" in cause_lower or "outside" in cause_lower:
        return "49 CFR 192.614 — Damage Prevention"
    if "pressure" in cause_lower or "incorrect" in cause_lower:
        return "49 CFR 192.605 — Operating Procedures"
    if "material" in cause_lower or "joint" in cause_lower:
        return "49 CFR 192.713 — Leak Surveys"
    if "natural" in cause_lower or "weather" in cause_lower:
        return "49 CFR 192.317 — Protection from Natural Hazards"
    if "weld" in cause_lower:
        return "49 CFR 192.719 — Pressure Testing"
    return "49 CFR 192.605 — General Operating Procedures"


def calculate_severity(row):
    try:
        cost = float(row.get("damage_cost", 0) or 0)
        fatal = float(row.get("fatalities", 0) or 0)
        if fatal > 0 or cost > 500000:
            return "Critical"
        if cost > 100000:
            return "High"
        if cost > 10000:
            return "Medium"
        return "Low"
    except Exception:
        return "Unknown"


def build_resolution_store(df):
    """Build structured resolution records from real PHMSA incidents."""
    records = []
    if df is None or df.empty:
        return records

    try:
        for index, row in df.iterrows():
            cause = str(row.get("cause", "") or "Unknown")
            city = str(row.get("city", "") or "Unknown")
            state = str(row.get("state", "") or "")
            records.append({
                "id": f"PHMSA_{index}",
                "employee_role": "Field Operations Engineer",
                "operator": str(row.get("operator", "") or ""),
                "issue_title": f"{cause} incident — {city} {state}".strip(),
                "issue_description": str(row.get("narrative", "") or ""),
                "symptoms": extract_symptoms(row.get("narrative", ""), cause),
                "root_cause": str(row.get("cause_details", "") or cause),
                "state": state,
                "year": row.get("year", ""),
                "city": city,
                "damage_cost": row.get("damage_cost", 0),
                "fatalities": row.get("fatalities", 0),
                "source": "PHMSA Federal Incident Database",
                "regulation": map_cause_to_regulation(cause),
                "severity": calculate_severity(row),
            })
    except Exception as exc:
        print(f"build_resolution_store failed: {exc}")
    return records


# ---------------------------------------------------------------------------
# PART 2 — RAG with ChromaDB
# ---------------------------------------------------------------------------

@st.cache_resource(show_spinner=False)
def get_embedder():
    from sentence_transformers import SentenceTransformer
    return SentenceTransformer("all-MiniLM-L6-v2")


def build_search_text(record):
    return f"""
Issue: {record['issue_title']}
Description: {record['issue_description']}
Root cause: {record['root_cause']}
Symptoms: {', '.join(record['symptoms'])}
Location: {record['city']} {record['state']}
Year: {record['year']}
Operator: {record['operator']}
Regulation: {record['regulation']}
Severity: {record['severity']}
""".strip()


@st.cache_resource(show_spinner=False)
def setup_chromadb(resolution_store_key, resolution_store):
    """resolution_store_key is a hashable proxy so cache invalidates when data changes."""
    import chromadb

    _ = resolution_store_key
    try:
        client = chromadb.Client()
        try:
            col = client.get_collection("mentor_resolutions")
            if col.count() > 0:
                return col
        except Exception:
            pass

        col = client.create_collection(
            name="mentor_resolutions",
            metadata={"hnsw:space": "cosine"},
        )

        if not resolution_store:
            return col

        embedder = get_embedder()
        batch_size = 100
        for i in range(0, len(resolution_store), batch_size):
            batch = resolution_store[i : i + batch_size]
            texts = [build_search_text(r) for r in batch]
            embeddings = embedder.encode(texts, show_progress_bar=False).tolist()
            col.add(
                documents=texts,
                embeddings=embeddings,
                ids=[r["id"] for r in batch],
                metadatas=[
                    {
                        "state": str(r["state"]),
                        "year": str(r["year"]),
                        "cause": str(r["root_cause"])[:100],
                        "severity": str(r["severity"]),
                        "operator": str(r["operator"])[:100],
                    }
                    for r in batch
                ],
            )
        return col
    except Exception as exc:
        print(f"setup_chromadb failed: {exc}")
        client = chromadb.Client()
        return client.create_collection(
            name="mentor_resolutions_fallback",
            metadata={"hnsw:space": "cosine"},
        )


def _keyword_search_resolutions(query, resolution_store, n=5):
    """Fallback text search when ChromaDB/embeddings are unavailable."""
    if not resolution_store:
        return []
    terms = [t.lower() for t in query.split() if len(t) > 2]
    scored = []
    for record in resolution_store:
        blob = " ".join(
            [
                str(record.get("issue_title", "")),
                str(record.get("issue_description", "")),
                str(record.get("root_cause", "")),
                " ".join(record.get("symptoms", [])),
                str(record.get("state", "")),
                str(record.get("city", "")),
            ]
        ).lower()
        hits = sum(1 for t in terms if t in blob)
        if hits or not terms:
            rec = dict(record)
            rec["relevance_score"] = round(min(hits / max(len(terms), 1), 1) * 100, 1)
            scored.append(rec)
    scored.sort(key=lambda r: r["relevance_score"], reverse=True)
    return scored[:n]


def search_resolutions(query, collection, resolution_store, n=5):
    try:
        if not resolution_store:
            return []
        if collection is None:
            return _keyword_search_resolutions(query, resolution_store, n=n)
        if collection.count() == 0:
            return _keyword_search_resolutions(query, resolution_store, n=n)

        embedder = get_embedder()
        q_emb = embedder.encode(query).tolist()
        results = collection.query(
            query_embeddings=[q_emb],
            n_results=min(n, collection.count()),
            include=["documents", "metadatas", "distances"],
        )

        matched_records = []
        for i, doc_id in enumerate(results["ids"][0]):
            record = next((r for r in resolution_store if r["id"] == doc_id), None)
            if record:
                rec = dict(record)
                rec["relevance_score"] = round(
                    (1 - results["distances"][0][i]) * 100, 1
                )
                matched_records.append(rec)
        return matched_records
    except Exception as exc:
        print(f"search_resolutions failed: {exc}")
        return _keyword_search_resolutions(query, resolution_store, n=n)


# ---------------------------------------------------------------------------
# PART 3 — Ollama answer generation
# ---------------------------------------------------------------------------

def generate_answer(query, matched_records):
    if not matched_records:
        return (
            "No matching incidents found in the Southern Company Gas incident database "
            "for this query. Try different keywords."
        )

    context = ""
    for i, r in enumerate(matched_records[:4]):
        context += f"""
INCIDENT {i + 1} [{r['state']}, {r['year']}]
Operator: {r['operator']}
Issue: {r['issue_title']}
What happened: {r['issue_description']}
Root cause: {r['root_cause']}
Symptoms: {', '.join(r['symptoms'])}
Regulation: {r['regulation']}
Severity: {r['severity']}
---"""

    system = f"""You are {APP_NAME} ({APP_FULL_NAME}), the institutional knowledge assistant for Southern Company. You answer questions using REAL incident data from the PHMSA federal pipeline safety database.

When answering:
1. Start with a direct 1-sentence answer
2. Reference the specific real incidents by state and year
3. Explain the root cause clearly
4. Give practical steps a field engineer would take
5. Cite the relevant CFR regulation
6. End with: VETERAN WATCH-OUT: [one thing only experienced engineers know]

Always say "According to real PHMSA incident data..." when referencing incidents.
Never make up data not in the context provided."""

    user = f"""Question: {query}

Real incident data from Southern Company territory:
{context}

Answer as {APP_NAME} — grounded in the real incidents above."""

    try:
        import ollama

        response = ollama.chat(
            model="llama3.2",
            messages=[
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
            options={"temperature": 0.2, "num_predict": 800},
        )
        return response["message"]["content"]
    except Exception as e:
        if "connect" in str(e).lower():
            return (
                "**Ollama not running.**\n\n"
                "Open terminal and run:\n"
                "```\nollama serve\n"
                "ollama pull llama3.2\n```\n"
                "Then refresh this page."
            )
        return f"Error: {str(e)}"


def _escape_html(text):
    return (
        str(text)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace("\n", "<br>")
    )


def _logo_data_uri():
    """Return a data URI for the Southern Company logo (PNG preferred, SVG fallback)."""
    for path, mime in (
        (LOGO_PNG_PATH, "image/png"),
        (LOGO_SVG_PATH, "image/svg+xml"),
    ):
        if os.path.isfile(path):
            encoded = base64.b64encode(open(path, "rb").read()).decode("ascii")
            return f"data:{mime};base64,{encoded}"
    return ""


def _render_header(incident_count):
    logo_uri = _logo_data_uri()
    logo_html = (
        f'<img src="{logo_uri}" alt="Southern Company" class="skim-logo-img" />'
        if logo_uri
        else "▲"
    )
    st.markdown(
        f"""
<div class="skim-header">
    <div class="skim-logo">{logo_html}</div>
    <div>
        <p class="skim-title">{APP_NAME}</p>
        <p class="skim-subtitle">{APP_FULL_NAME}</p>
        <p class="skim-tagline">{APP_TAGLINE}</p>
    </div>
    <div class="skim-badge">{incident_count:,} real incidents · Live PHMSA data</div>
</div>
""",
        unsafe_allow_html=True,
    )


def _load_mentor_data():
    """Load all data sources with per-step error handling; never raises."""
    empty_incidents = _empty_incidents_df()
    empty_annual = pd.DataFrame(
        columns=["year", "operator", "pipeline_miles", "services", "surveyed"]
    )
    empty_sec = pd.DataFrame(columns=["filing_date", "form", "description", "accession"])
    empty_ferc = pd.DataFrame()
    errors = []

    df = empty_incidents
    phmsa_errors = []
    try:
        df, phmsa_errors = fetch_phmsa_incidents()
    except Exception as exc:
        phmsa_errors = [str(exc)]
        errors.append(f"PHMSA incidents: {exc}")

    annual_stats = empty_annual
    try:
        annual_stats = fetch_phmsa_annual_stats()
    except Exception as exc:
        errors.append(f"PHMSA annual stats: {exc}")

    sec_filings = empty_sec
    try:
        sec_filings = fetch_sec_filings()
    except Exception as exc:
        errors.append(f"SEC filings: {exc}")

    ferc_data = empty_ferc
    try:
        ferc_data = fetch_ferc_gas_data()
    except Exception as exc:
        errors.append(f"FERC data: {exc}")

    resolution_store = []
    try:
        resolution_store = build_resolution_store(df)
    except Exception as exc:
        errors.append(f"Resolution store: {exc}")

    collection = None
    if resolution_store:
        try:
            collection = setup_chromadb(len(resolution_store), resolution_store)
        except Exception as exc:
            errors.append(f"Semantic search: {exc}")

    return {
        "df": df,
        "annual_stats": annual_stats,
        "sec_filings": sec_filings,
        "ferc_data": ferc_data,
        "resolution_store": resolution_store,
        "collection": collection,
        "errors": errors,
        "phmsa_errors": phmsa_errors,
    }


# ---------------------------------------------------------------------------
# PART 4 — Dashboard UI
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title=f"{APP_NAME} · {APP_FULL_NAME}",
    layout="wide",
    page_icon=LOGO_PNG_PATH if os.path.isfile(LOGO_PNG_PATH) else "🔺",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
<style>
.block-container { padding-top: 1.5rem; padding-bottom: 2rem; }
.skim-header {
    display: flex; align-items: center; gap: 16px;
    padding: 1.25rem 1.5rem; background: #0A2E52;
    border-radius: 14px; margin-bottom: 1.5rem;
}
.skim-logo {
    width: 56px; height: 56px; background: white;
    border-radius: 10px; display: flex; align-items: center;
    justify-content: center; font-size: 22px; flex-shrink: 0;
    padding: 6px; box-shadow: 0 1px 4px rgba(0,0,0,0.15);
}
.skim-logo-img { width: 100%; height: 100%; object-fit: contain; display: block; }
.skim-title { color: white; font-size: 28px; font-weight: 700; margin: 0; letter-spacing: 0.08em; }
.skim-subtitle { color: white; font-size: 15px; font-weight: 500; margin: 2px 0 0; }
.skim-tagline { color: #7EB3E8; font-size: 12px; margin: 4px 0 0; }
.skim-badge {
    margin-left: auto; background: rgba(255,255,255,0.1); color: #7EB3E8;
    padding: 6px 14px; border-radius: 20px; font-size: 12px;
    border: 1px solid rgba(255,255,255,0.15); flex-shrink: 0;
}
.answer-card {
    background: white; border: 1px solid #E2E8F0;
    border-left: 5px solid #185FA5; border-radius: 0 14px 14px 0;
    padding: 1.5rem 1.75rem; font-size: 15px; line-height: 1.85;
    color: #1A202C; margin: 1rem 0;
}
.watch-out-box {
    background: #FFFBEB; border: 1px solid #F6E05E;
    border-radius: 10px; padding: 12px 16px; margin-top: 1rem;
    font-size: 13px; color: #744210;
}
.incident-card {
    background: white; border: 1px solid #E2E8F0; border-radius: 12px;
    padding: 1rem 1.25rem; margin-bottom: 10px; transition: border-color 0.15s;
}
.incident-card:hover { border-color: #185FA5; }
.incident-header {
    display: flex; justify-content: space-between;
    align-items: flex-start; margin-bottom: 8px;
}
.incident-title { font-size: 13px; font-weight: 600; color: #1A202C; flex: 1; }
.severity-badge {
    font-size: 11px; padding: 3px 10px; border-radius: 20px;
    font-weight: 500; margin-left: 8px; flex-shrink: 0;
}
.sev-critical { background:#FED7D7; color:#822727; }
.sev-high { background:#FEEBC8; color:#7B341E; }
.sev-medium { background:#FEFCBF; color:#744210; }
.sev-low { background:#C6F6D5; color:#22543D; }
.incident-meta { font-size: 12px; color: #718096; margin-bottom: 6px; }
.incident-narrative {
    font-size: 12px; color: #4A5568; line-height: 1.6;
    display: -webkit-box; -webkit-line-clamp: 3;
    -webkit-box-orient: vertical; overflow: hidden;
}
.relevance-bar {
    height: 3px; background: #E2E8F0; border-radius: 3px;
    margin-top: 10px; overflow: hidden;
}
.relevance-fill { height: 100%; background: #185FA5; border-radius: 3px; }
.status-dot-green {
    display: inline-block; width: 8px; height: 8px;
    background: #38A169; border-radius: 50%; margin-right: 6px;
}
.status-dot-red {
    display: inline-block; width: 8px; height: 8px;
    background: #E53E3E; border-radius: 50%; margin-right: 6px;
}
#MainMenu { visibility: hidden; }
footer { visibility: hidden; }
.stDeployButton { display: none; }
[data-testid="baseButton-primary"] {
    background: #185FA5 !important; color: white !important;
    border: none !important; font-size: 15px !important;
    padding: 12px 24px !important; border-radius: 10px !important;
    font-weight: 600 !important;
}
[data-testid="baseButton-primary"]:hover { background: #0C447C !important; }
</style>
""",
    unsafe_allow_html=True,
)

if "history" not in st.session_state:
    st.session_state.history = []
if "active_query" not in st.session_state:
    st.session_state.active_query = ""
if "last_results" not in st.session_state:
    st.session_state.last_results = []
if "last_answer" not in st.session_state:
    st.session_state.last_answer = ""
if "mentor_data" not in st.session_state:
    st.session_state.mentor_data = None

# Render header immediately so the page is never blank
_render_header(
    len(st.session_state.mentor_data["resolution_store"])
    if st.session_state.mentor_data
    else 0
)

if st.session_state.mentor_data is None:
    load_errors = []
    with st.status("Connecting to real Southern Company data...", expanded=True) as status:
        st.write("Fetching PHMSA federal incident database...")
        try:
            st.session_state.mentor_data = _load_mentor_data()
            load_errors = st.session_state.mentor_data.get("errors", [])
            md = st.session_state.mentor_data
            st.write(f"Loaded {len(md['df']):,} real incidents from Southern Company territory")
            st.write(f"Loaded {len(md['annual_stats']):,} annual report records")
            st.write(f"Loaded {len(md['sec_filings']):,} SEC filing records")
            st.write(f"Loaded {len(md['ferc_data']):,} FERC data rows")
            st.write(f"Indexed {len(md['resolution_store']):,} incident resolutions")
            if md["collection"] is not None:
                st.write("Semantic search ready")
            else:
                st.write("Semantic search unavailable — text search fallback active")
            if load_errors:
                st.warning("Some data sources failed: " + "; ".join(load_errors[:3]))
            if status is not None:
                status.update(
                    label=(
                        f"{APP_NAME} ready · {len(md['resolution_store']):,} real incidents indexed"
                        if md["resolution_store"]
                        else f"{APP_NAME} ready · limited data (check connection)"
                    ),
                    state="complete",
                    expanded=bool(load_errors),
                )
        except Exception as exc:
            st.session_state.mentor_data = {
                "df": _empty_incidents_df(),
                "annual_stats": pd.DataFrame(),
                "sec_filings": pd.DataFrame(),
                "ferc_data": pd.DataFrame(),
                "resolution_store": [],
                "collection": None,
                "errors": [str(exc)],
                "phmsa_errors": [str(exc)],
            }
            st.error(f"Startup error: {exc}")
            if status is not None:
                status.update(label=f"{APP_NAME} started with errors", state="error", expanded=True)

data = st.session_state.mentor_data
df = data["df"]
annual_stats = data["annual_stats"]
sec_filings = data["sec_filings"]
ferc_data = data["ferc_data"]
resolution_store = data["resolution_store"]
collection = data["collection"]
load_errors = data.get("errors", [])
phmsa_fetch_errors = data.get("phmsa_errors", [])

if load_errors:
    with st.expander("Data source warnings", expanded=False):
        for err in load_errors:
            st.caption(f"⚠ {err}")

left_col, right_col = st.columns([5, 5], gap="large")

with left_col:
    st.markdown(f"#### Search {APP_NAME}")
    st.caption(
        "Ask anything in plain English about gas operations, incidents, safety procedures, "
        "or regulations. Results come from real Southern Company incident data."
    )

    suggestions = [
        "What causes gas leaks in distribution mains?",
        "How are pipeline pressure drops resolved?",
        "What happens during excavation damage incidents?",
        "Show me corrosion incidents in Georgia",
        "What are the most severe incidents in Alabama?",
        "How long do pipeline repairs typically take?",
        "What regulations apply to leak detection?",
        "Show me incidents from cold weather events",
    ]

    st.markdown(
        "<p style='font-size:12px;color:#718096;margin-bottom:6px;font-weight:500'>"
        "Quick searches</p>",
        unsafe_allow_html=True,
    )

    col1, col2 = st.columns(2)
    for i, s in enumerate(suggestions):
        target_col = col1 if i % 2 == 0 else col2
        with target_col:
            if st.button(s, key=f"sug_{i}", use_container_width=True):
                st.session_state.active_query = s

    st.markdown("<br>", unsafe_allow_html=True)

    query_input = st.text_area(
        "Your question:",
        value=st.session_state.active_query,
        placeholder=(
            "Type your question in plain English...\n\n"
            "Examples:\n"
            "• What caused the most damage in Tennessee?\n"
            "• How do field crews handle gas odor reports?\n"
            "• What are common failure modes in winter?"
        ),
        height=120,
        label_visibility="collapsed",
    )

    col_btn, col_clear = st.columns([3, 1])
    with col_btn:
        search_btn = st.button(f"Search {APP_NAME} →", type="primary", use_container_width=True)
    with col_clear:
        if st.button("Clear", use_container_width=True):
            st.session_state.active_query = ""
            st.session_state.last_results = []
            st.session_state.last_answer = ""
            st.rerun()

    if st.session_state.history:
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(
            "<p style='font-size:12px;color:#718096;font-weight:500'>Recent searches</p>",
            unsafe_allow_html=True,
        )
        for idx, h in enumerate(reversed(st.session_state.history[-6:])):
            label = f"↩ {h[:55]}..." if len(h) > 55 else f"↩ {h}"
            if st.button(label, key=f"hist_{idx}_{hash(h)}", use_container_width=True):
                st.session_state.active_query = h
                st.rerun()

with right_col:
    st.markdown("#### Southern Company Incident Data")

    if len(df) > 0:
        m1, m2, m3, m4 = st.columns(4)

        total = len(df)
        states = df["state"].nunique() if "state" in df.columns else 0
        years = df["year"].nunique() if "year" in df.columns else 0
        try:
            total_damage = df["damage_cost"].sum()
            damage_str = f"${total_damage / 1e6:.1f}M"
        except Exception:
            damage_str = "N/A"

        m1.metric("Real Incidents", f"{total:,}")
        m2.metric("States", states)
        m3.metric("Years Covered", years)
        m4.metric("Total Damage", damage_str)

        tab1, tab2, tab3 = st.tabs(["By Cause", "By Year", "By State"])

        with tab1:
            if "cause" in df.columns:
                cause_df = (
                    df["cause"].str.strip().str.title().value_counts().head(8).reset_index()
                )
                cause_df.columns = ["Cause", "Incidents"]
                cause_df["Cause"] = cause_df["Cause"].str[:40]
                fig = px.bar(
                    cause_df,
                    x="Incidents",
                    y="Cause",
                    orientation="h",
                    color="Incidents",
                    color_continuous_scale=["#C7DEFF", "#185FA5"],
                )
                fig.update_layout(
                    height=260,
                    margin=dict(l=0, r=0, t=10, b=0),
                    plot_bgcolor="rgba(0,0,0,0)",
                    paper_bgcolor="rgba(0,0,0,0)",
                    coloraxis_showscale=False,
                    yaxis=dict(autorange="reversed"),
                    font=dict(size=11),
                    yaxis_title="",
                    xaxis_title="Number of incidents",
                )
                fig.update_xaxes(showgrid=True, gridcolor="rgba(0,0,0,0.05)")
                fig.update_yaxes(showgrid=False)
                st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

        with tab2:
            if "year" in df.columns:
                year_df = df.copy()
                year_df["year"] = pd.to_numeric(year_df["year"], errors="coerce")
                year_df = (
                    year_df.groupby("year")
                    .agg(Incidents=("year", "count"), Damage=("damage_cost", "sum"))
                    .reset_index()
                    .dropna(subset=["year"])
                )
                year_df["year"] = year_df["year"].astype(int)

                fig2 = px.bar(
                    year_df,
                    x="year",
                    y="Incidents",
                    color="Incidents",
                    color_continuous_scale=["#C7DEFF", "#0A2E52"],
                )
                fig2.update_layout(
                    height=260,
                    margin=dict(l=0, r=0, t=10, b=0),
                    plot_bgcolor="rgba(0,0,0,0)",
                    paper_bgcolor="rgba(0,0,0,0)",
                    coloraxis_showscale=False,
                    font=dict(size=11),
                    xaxis_title="Year",
                    yaxis_title="Incidents",
                )
                fig2.update_xaxes(showgrid=False)
                fig2.update_yaxes(showgrid=True, gridcolor="rgba(0,0,0,0.05)")
                st.plotly_chart(fig2, use_container_width=True, config={"displayModeBar": False})

        with tab3:
            if "state" in df.columns:
                state_df = df["state"].value_counts().reset_index()
                state_df.columns = ["State", "Incidents"]

                fig3 = px.bar(
                    state_df,
                    x="State",
                    y="Incidents",
                    color="Incidents",
                    color_continuous_scale=["#C7DEFF", "#0A2E52"],
                    text="Incidents",
                )
                fig3.update_traces(textposition="outside")
                fig3.update_layout(
                    height=260,
                    margin=dict(l=0, r=0, t=10, b=30),
                    plot_bgcolor="rgba(0,0,0,0)",
                    paper_bgcolor="rgba(0,0,0,0)",
                    coloraxis_showscale=False,
                    font=dict(size=11),
                    xaxis_title="",
                    yaxis_title="Incidents",
                )
                fig3.update_xaxes(showgrid=False)
                fig3.update_yaxes(showgrid=True, gridcolor="rgba(0,0,0,0.05)")
                st.plotly_chart(fig3, use_container_width=True, config={"displayModeBar": False})
    else:
        if phmsa_fetch_errors:
            st.warning("PHMSA fetch errors:\n\n" + "\n".join(f"• {e}" for e in phmsa_fetch_errors[:4]))
        else:
            st.info("No live PHMSA incidents loaded. See sidebar for fetch details.")

active_q = query_input or st.session_state.active_query

if (search_btn or st.session_state.active_query) and active_q and active_q.strip():
    if active_q not in st.session_state.history:
        st.session_state.history.append(active_q)

    st.divider()
    st.markdown(f"### Results for: *{active_q}*")

    with st.spinner("Searching real Southern Company incident data..."):
        matched = search_resolutions(active_q, collection, resolution_store, n=6)
        st.session_state.last_results = matched

    answer_col, incidents_col = st.columns([5, 5], gap="large")

    with answer_col:
        st.markdown(f"#### {APP_NAME}'s Answer")
        st.caption("Generated from real PHMSA incident data using semantic search + Ollama llama3.2")

        with st.spinner("Generating expert answer..."):
            answer = generate_answer(active_q, matched)
            st.session_state.last_answer = answer

        answer_html = _escape_html(answer)
        if "watch-out" in answer.lower() or "veteran" in answer.lower():
            lower = answer.lower()
            idx = lower.find("watch-out:")
            if idx >= 0:
                main_answer = _escape_html(answer[:idx])
                watch_out = _escape_html(answer[idx:])
                st.markdown(f"<div class='answer-card'>{main_answer}</div>", unsafe_allow_html=True)
                st.markdown(f"<div class='watch-out-box'>⚠ {watch_out}</div>", unsafe_allow_html=True)
            else:
                st.markdown(f"<div class='answer-card'>{answer_html}</div>", unsafe_allow_html=True)
        else:
            st.markdown(f"<div class='answer-card'>{answer_html}</div>", unsafe_allow_html=True)

    with incidents_col:
        st.markdown(f"#### {len(matched)} Matching Real Incidents")
        st.caption("Ranked by semantic relevance to your query. All from PHMSA federal database.")

        if matched:
            for r in matched:
                sev = r.get("severity", "Unknown")
                sev_class = {
                    "Critical": "sev-critical",
                    "High": "sev-high",
                    "Medium": "sev-medium",
                    "Low": "sev-low",
                }.get(sev, "sev-low")

                relevance = r.get("relevance_score", 0)
                narrative = str(r.get("issue_description", ""))
                short_narrative = narrative[:200] + "..." if len(narrative) > 200 else narrative
                reg = _escape_html(r.get("regulation", ""))
                symptoms = _escape_html(", ".join(r.get("symptoms", [])))

                st.markdown(
                    f"""
<div class='incident-card'>
  <div class='incident-header'>
    <div class='incident-title'>{_escape_html(r.get('issue_title', 'Unknown incident'))}</div>
    <span class='severity-badge {sev_class}'>{sev}</span>
  </div>
  <div class='incident-meta'>
    📍 {_escape_html(r.get('city',''))} {_escape_html(r.get('state',''))} ·
    📅 {_escape_html(r.get('year',''))} ·
    🏢 {_escape_html(str(r.get('operator',''))[:30])}
  </div>
  <div class='incident-meta'>⚠ Root cause: {_escape_html(r.get('root_cause',''))}</div>
  <div class='incident-meta'>🔍 Symptoms: {symptoms[:80]}</div>
  <div class='incident-narrative'>{_escape_html(short_narrative)}</div>
  <div class='incident-meta' style='margin-top:8px;color:#185FA5'>📋 {reg}</div>
  <div class='relevance-bar'>
    <div class='relevance-fill' style='width:{relevance}%'></div>
  </div>
  <div style='font-size:10px;color:#A0AEC0;text-align:right;margin-top:2px'>
    {relevance}% relevance match
  </div>
</div>""",
                    unsafe_allow_html=True,
                )
        else:
            st.info("No matching incidents found. Try different keywords or a broader query.")

    st.session_state.active_query = ""

with st.sidebar:
    if os.path.isfile(LOGO_PNG_PATH):
        st.image(LOGO_PNG_PATH, width=72)
    elif os.path.isfile(LOGO_SVG_PATH):
        st.image(LOGO_SVG_PATH, width=72)
    st.markdown(f"### {APP_NAME}")
    st.caption(APP_FULL_NAME)
    st.caption(APP_TAGLINE)

    st.divider()
    st.markdown("**System Status**")

    try:
        import ollama

        models = ollama.list()
        st.markdown(
            "<span class='status-dot-green'></span>Ollama running",
            unsafe_allow_html=True,
        )
        model_names = [m.get("name", "") for m in models.get("models", [])]
        if any("llama3.2" in n for n in model_names):
            st.caption("✓ llama3.2 ready")
        else:
            st.caption("⚠ Run: ollama pull llama3.2")
    except Exception:
        st.markdown(
            "<span class='status-dot-red'></span>Ollama offline",
            unsafe_allow_html=True,
        )
        st.code("ollama serve\nollama pull llama3.2", language="bash")

    st.divider()
    st.markdown("**Data Sources**")

    phmsa_ok = len(df) > 0
    dot = "green" if phmsa_ok else "red"
    st.markdown(
        f"<span class='status-dot-{dot}'></span>PHMSA Federal Database",
        unsafe_allow_html=True,
    )
    if not phmsa_ok and phmsa_fetch_errors:
        retired_note = next((e for e in phmsa_fetch_errors if "retired" in e.lower()), None)
        if retired_note:
            st.warning(retired_note)
        st.error("PHMSA incidents blocked or empty:\n\n" + "\n".join(
            f"• {e}" for e in phmsa_fetch_errors[:6] if e != retired_note
        ))
        st.caption(
            "Tip: if .gov downloads are blocked, download PHMSA data in your browser to "
            "`~/Downloads/PHMSA_Pipeline_Safety_Flagged_Incidents.zip` — SKIM loads it automatically."
        )
    st.caption(
        f"{len(resolution_store):,} real incidents\n\n"
        f"States: GA, AL, TN, MS, FL, VA, NC\n\n"
        f"Live data · Updates hourly"
    )

    annual_ok = len(annual_stats) > 0
    dot2 = "green" if annual_ok else "red"
    st.markdown(
        f"<span class='status-dot-{dot2}'></span>PHMSA Annual Reports",
        unsafe_allow_html=True,
    )
    if annual_ok:
        latest = annual_stats.iloc[0]
        st.caption(
            f"{len(annual_stats)} records · Latest: {latest.get('year', 'N/A')}\n\n"
            f"Pipeline miles: {latest.get('pipeline_miles', 'N/A'):,.0f}"
            if pd.notna(latest.get("pipeline_miles"))
            else f"{len(annual_stats)} annual report records"
        )
    else:
        st.caption("Annual stats unavailable")

    sec_ok = len(sec_filings) > 0
    dot3 = "green" if sec_ok else "red"
    st.markdown(
        f"<span class='status-dot-{dot3}'></span>SEC EDGAR Filings",
        unsafe_allow_html=True,
    )
    if sec_ok:
        st.caption(f"{len(sec_filings)} recent Southern Company filings")
    else:
        st.caption("SEC filings unavailable")

    ferc_ok = len(ferc_data) > 0
    dot4 = "green" if ferc_ok else "red"
    st.markdown(
        f"<span class='status-dot-{dot4}'></span>FERC Gas Data",
        unsafe_allow_html=True,
    )
    st.caption(
        f"{len(ferc_data)} pipeline capacity rows"
        if ferc_ok
        else "FERC data unavailable"
    )

    st.divider()
    st.markdown("**How it works**")
    st.caption(
        "1. You type any question in plain English\n\n"
        f"2. {APP_NAME} searches real PHMSA incident data using semantic similarity\n\n"
        "3. Most relevant real incidents are retrieved\n\n"
        "4. Ollama generates an expert answer grounded in those real incidents\n\n"
        "5. You see both the answer AND the raw incident records that support it"
    )

    st.divider()
    st.caption(
        "Cox Play With Purpose Hackathon 2026\n"
        "Track 01: Energy\n"
        "Problem 5: Retirement Transition\n"
        "Southern Company · New Ventures"
    )
