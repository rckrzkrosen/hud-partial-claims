#!/usr/bin/env python3
"""
HUD Partial Claim Refresh — NYC (all 5 boroughs)
==================================================
Run this once a day (same schedule as acris_monitor.py). Every run does a
full rebuild, so it is always both the complete list AND up to date — no
separate "backfill mode" needed.

WHAT IT PRODUCES (3 files, written next to this script)
----------------------------------------------------------
  hud_partial_claims.json   <- upload this next to hud-search.html on the
                                website. This is what the search page reads.
  hud_partial_claims.csv    <- same data, for opening in Excel/Numbers.
  seen_ids_hud_partial_claim.json  <- internal bookkeeping only, do not
                                upload this one. Used only to know which
                                rows are new since the last run, for the
                                email.

WHAT COUNTS AS A "HUD PARTIAL CLAIM"
----------------------------------------------------------
ACRIS has no dedicated document type for this. It's an ordinary mortgage
document where the lender (party_type "2") is the U.S. Department of
Housing and Urban Development, recorded when an FHA-insured borrower fell
behind and HUD covered the arrears with a deferred, interest-free second
lien. This script finds every one of those, citywide, and resolves each to
a street address and zip code.

DATA SOURCES — NYC Open Data / Socrata, all public, no key required
----------------------------------------------------------------------
  ACRIS - Real Property Parties : 636b-3b5g  (who's the lender, and
                                                who's the borrower — same
                                                dataset, party_type '1'
                                                vs '2')
  ACRIS - Real Property Master  : bnx9-e6tj  (the document itself)
  ACRIS - Real Property Legals  : 8h5j-fqxa  (which property, block/lot)
  PLUTO (property database)     : 64uk-42ks  (block/lot -> zip code)

OWNER NAME + ABSENTEE-OWNER FLAG (added Sept 2026)
----------------------------------------------------------------------
The same Real Property Parties dataset already being queried for "who's
the lender" also has a row for "who's the borrower" (party_type '1'),
and that row includes the borrower's own mailing address — confirmed
live: for document 2026082600319001, the party_type '1' row is
"MGBEME, LYNNE" at "10921 SPRINGFIELD BLVD, QUEENS VILLAGE, NY 11429",
matching the property's own address exactly.

So for every HUD Partial Claim record, this script now also pulls the
borrower's name and mailing zip code, and compares that mailing zip to
the property's own zip code (from PLUTO). A mismatch means the mail for
this loan doesn't go to the property itself — the classic sign of an
absentee/out-of-town landlord rather than an owner-occupant. This adds
no new dataset and no new daily job — it's one more thing read off data
already being fetched.

USAGE
-----
    pip install requests
    python3 hud_partial_claim_refresh.py

Optional email (same env vars acris_monitor.py already uses, if any —
reuse them here instead of setting up a second set):
    ACRIS_EMAIL_TO      default: rrazack@rosenjacob.com
    ACRIS_EMAIL_FROM
    ACRIS_SMTP_HOST
    ACRIS_SMTP_PORT      default: 587
    ACRIS_SMTP_USER
    ACRIS_SMTP_PASS

If those aren't set, the script still runs fine — it just skips the email
and prints a summary instead. The JSON/CSV files are written either way.
"""

import csv
import json
import os
import smtplib
import sys
import time
from email.mime.text import MIMEText
from pathlib import Path
from urllib.parse import urlencode

import requests

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

SOCRATA_BASE = "https://data.cityofnewyork.us/resource"
PARTIES_DATASET = "636b-3b5g"
MASTER_DATASET = "bnx9-e6tj"
LEGALS_DATASET = "8h5j-fqxa"
PLUTO_DATASET = "64uk-42ks"

# ACRIS Legals uses numeric borough codes; PLUTO uses 2-letter codes.
BOROUGH_CODE_TO_NAME = {
    "1": "Manhattan", "2": "Bronx", "3": "Brooklyn",
    "4": "Queens", "5": "Staten Island",
}
BOROUGH_CODE_TO_PLUTO = {
    "1": "MN", "2": "BX", "3": "BK", "4": "QN", "5": "SI",
}

# Mortgage-family document types worth keeping (excludes unrelated HUD
# paperwork that happens to name HUD as a party but isn't a lien).
RELEVANT_DOC_TYPES = {"MTGE", "M&CON", "CORRM", "MMTG", "SMTG", "CMTG"}

STATE_DIR = Path(__file__).resolve().parent
SEEN_IDS_PATH = STATE_DIR / "seen_ids_hud_partial_claim.json"
JSON_OUTPUT_PATH = STATE_DIR / "hud_partial_claims.json"
CSV_OUTPUT_PATH = STATE_DIR / "hud_partial_claims.csv"

PAGE_SIZE = 1000
REQUEST_PAUSE_SEC = 0.2
# How many document_ids go into one "in (...)" lookup. NYC's server 400s on
# a web address that gets too long. 200 at a time was too many (several
# thousand characters); POSTing instead avoided the length problem but NYC's
# server turned out to require a login for POST requests specifically, which
# is worse. 40 at a time keeps the plain web address short enough to always
# be accepted, with no login needed — just more, smaller requests instead.
BATCH_SIZE = 40


def get_json(dataset: str, params: dict, max_retries: int = 5) -> list:
    """Fetch one page from NYC Open Data, retrying on slow/dropped connections.

    NYC's server occasionally takes longer than 30 seconds to answer,
    especially on the "name like %...%" search this script starts with.
    That used to crash the whole run on attempt #1. Now it waits longer,
    and if a request still fails, it tries again a few times (5s, 10s,
    20s... between tries) before giving up on that one page.
    """
    url = f"{SOCRATA_BASE}/{dataset}.json?{urlencode(params)}"
    last_error = None
    for attempt in range(1, max_retries + 1):
        try:
            resp = requests.get(url, timeout=90)
            resp.raise_for_status()
            return resp.json()
        except (requests.exceptions.Timeout,
                requests.exceptions.ConnectionError) as e:
            last_error = e
            if attempt < max_retries:
                wait = 5 * (2 ** (attempt - 1))  # 5, 10, 20, 40...
                print(f"    (network hiccup, retrying in {wait}s — "
                      f"attempt {attempt}/{max_retries})")
                time.sleep(wait)
            continue
        except requests.exceptions.HTTPError as e:
            status = e.response.status_code if e.response is not None else None
            body = ""
            try:
                body = e.response.text[:300] if e.response is not None else ""
            except Exception:
                pass
            last_error = e
            if status in (429, 500, 502, 503, 504) and attempt < max_retries:
                wait = 5 * (2 ** (attempt - 1))
                print(f"    (server returned {status}, retrying in {wait}s — "
                      f"attempt {attempt}/{max_retries})")
                time.sleep(wait)
                continue
            if body:
                print(f"    NYC Open Data said: {body}")
            raise
    raise last_error


def chunked(seq, size):
    seq = list(seq)
    for i in range(0, len(seq), size):
        yield seq[i:i + size]


def fetch_hud_partial_claim_document_ids() -> list:
    """All document_ids where HUD is the lender (party_type '2')."""
    doc_ids = set()
    offset = 0
    where_clause = (
        "party_type = '2' AND "
        "(upper(name) like '%SECRETARY OF HOUSING AND URBAN DEVELOPMENT%')"
    )
    while True:
        rows = get_json(PARTIES_DATASET, {
            "$select": "document_id",
            "$where": where_clause,
            "$limit": PAGE_SIZE,
            "$offset": offset,
        })
        if not rows:
            break
        doc_ids.update(r["document_id"] for r in rows)
        print(f"  ...parties offset {offset}: {len(rows)} rows "
              f"({len(doc_ids)} unique document_ids so far)")
        if len(rows) < PAGE_SIZE:
            break
        offset += PAGE_SIZE
        time.sleep(REQUEST_PAUSE_SEC)
    return sorted(doc_ids)


def fetch_master_info(document_ids: list) -> dict:
    """document_id -> {doc_type, doc_amount, recorded_datetime}.

    NYC's real column name for the dollar amount on this dataset is
    "document_amt", not "doc_amount" — confirmed by querying the live
    dataset directly (asking for "doc_amount" gets a
    query.soql.no-such-column error). We ask for the real name, then
    rename it to "doc_amount" right away so nothing downstream has to
    change.
    """
    info = {}
    for batch in chunked(document_ids, BATCH_SIZE):
        id_list = ",".join(f"'{d}'" for d in batch)
        rows = get_json(MASTER_DATASET, {
            "$select": "document_id, doc_type, document_amt, recorded_datetime",
            "$where": f"document_id in ({id_list})",
            "$limit": BATCH_SIZE,
        })
        for row in rows:
            if row.get("doc_type") in RELEVANT_DOC_TYPES:
                row["doc_amount"] = row.pop("document_amt", "")
                info[row["document_id"]] = row
        time.sleep(REQUEST_PAUSE_SEC)
    return info


def fetch_borrower_info(document_ids: list) -> dict:
    """document_id -> {name, zip, city, state} for the borrower (party_type '1').

    A document can have more than one party_type '1' row (co-borrowers).
    We keep the first one — for the absentee-owner comparison, that's
    plenty: if the primary borrower's mail doesn't go to the property,
    that's the signal we care about.
    """
    info = {}
    for batch in chunked(document_ids, BATCH_SIZE):
        id_list = ",".join(f"'{d}'" for d in batch)
        rows = get_json(PARTIES_DATASET, {
            "$select": "document_id, name, zip, city, state",
            "$where": f"document_id in ({id_list}) AND party_type = '1'",
            "$limit": BATCH_SIZE * 2,
        })
        for row in rows:
            doc_id = row["document_id"]
            if doc_id not in info:
                info[doc_id] = row
        time.sleep(REQUEST_PAUSE_SEC)
    return info


def fetch_legals_info(document_ids: list) -> dict:
    """document_id -> list of {borough, block, lot, street_number, street_name, unit}."""
    info = {}
    for batch in chunked(document_ids, BATCH_SIZE):
        id_list = ",".join(f"'{d}'" for d in batch)
        rows = get_json(LEGALS_DATASET, {
            "$select": "document_id, borough, block, lot, street_number, "
                       "street_name, unit",
            "$where": f"document_id in ({id_list})",
            "$limit": BATCH_SIZE * 3,
        })
        for row in rows:
            info.setdefault(row["document_id"], []).append(row)
        time.sleep(REQUEST_PAUSE_SEC)
    return info


def make_bbl(borough_code: str, block: str, lot: str) -> str:
    try:
        return f"{int(borough_code)}{int(block):05d}{int(lot):04d}"
    except (ValueError, TypeError):
        return ""


def fetch_zip_codes(bbls: list) -> dict:
    """bbl -> zipcode, via PLUTO."""
    info = {}
    clean_bbls = [b for b in set(bbls) if b]
    for batch in chunked(clean_bbls, BATCH_SIZE):
        id_list = ",".join(f"'{b}'" for b in batch)
        rows = get_json(PLUTO_DATASET, {
            "$select": "bbl, zipcode",
            "$where": f"bbl in ({id_list})",
            "$limit": BATCH_SIZE,
        })
        for row in rows:
            bbl_clean = row["bbl"].split(".")[0]  # PLUTO pads with .00000000
            info[bbl_clean] = row.get("zipcode", "")
        time.sleep(REQUEST_PAUSE_SEC)
    return info


def build_rows(master_info: dict, legals_info: dict, borrower_info: dict) -> list:
    rows = []
    all_bbls = []
    prelim = []
    for doc_id, master in master_info.items():
        for legal in legals_info.get(doc_id, [{}]):
            borough_code = legal.get("borough", "")
            bbl = make_bbl(borough_code, legal.get("block", "0"), legal.get("lot", "0"))
            all_bbls.append(bbl)
            prelim.append((doc_id, master, legal, bbl))

    zip_lookup = fetch_zip_codes(all_bbls)

    for doc_id, master, legal, bbl in prelim:
        borough_code = legal.get("borough", "")
        street_num = (legal.get("street_number") or "").strip()
        street_name = (legal.get("street_name") or "").strip()
        unit = (legal.get("unit") or "").strip()
        address = " ".join(p for p in [street_num, street_name] if p)
        if unit:
            address = f"{address}, Unit {unit}"

        property_zip = zip_lookup.get(bbl, "")
        borrower = borrower_info.get(doc_id, {})
        owner_name = (borrower.get("name") or "").strip()
        owner_zip = (borrower.get("zip") or "").strip()[:5]
        owner_city = (borrower.get("city") or "").strip()
        owner_state = (borrower.get("state") or "").strip()

        # Absentee = the borrower's own mail doesn't go to this property.
        # Only call it when we actually have both zips to compare — an
        # unknown mailing address is not the same as a confirmed mismatch.
        absentee_owner = bool(
            property_zip and owner_zip and property_zip[:5] != owner_zip
        )

        rows.append({
            "document_id": doc_id,
            "borough": BOROUGH_CODE_TO_NAME.get(borough_code, borough_code),
            "block": legal.get("block", ""),
            "lot": legal.get("lot", ""),
            "zip": property_zip,
            "address": address,
            "recorded_date": (master.get("recorded_datetime") or "")[:10],
            "doc_amount": master.get("doc_amount", ""),
            "doc_type": master.get("doc_type", ""),
            "owner_name": owner_name,
            "owner_mailing_city_state": (
                ", ".join(p for p in [owner_city, owner_state] if p)
            ),
            "absentee_owner": absentee_owner,
        })
    return rows


def load_seen_ids() -> set:
    if SEEN_IDS_PATH.exists():
        return set(json.loads(SEEN_IDS_PATH.read_text()))
    return set()


def save_seen_ids(ids: set) -> None:
    SEEN_IDS_PATH.write_text(json.dumps(sorted(ids), indent=2))


def write_outputs(rows: list) -> None:
    rows_sorted = sorted(rows, key=lambda r: r["recorded_date"], reverse=True)

    JSON_OUTPUT_PATH.write_text(json.dumps({
        "generated_at": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "count": len(rows_sorted),
        "properties": rows_sorted,
    }, indent=2))

    fieldnames = ["document_id", "borough", "zip", "block", "lot", "address",
                  "recorded_date", "doc_amount", "doc_type", "owner_name",
                  "owner_mailing_city_state", "absentee_owner"]
    with open(CSV_OUTPUT_PATH, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows_sorted)


def send_email(new_rows: list) -> None:
    to_addr = os.environ.get("ACRIS_EMAIL_TO", "rrazack@rosenjacob.com")
    from_addr = os.environ.get("ACRIS_EMAIL_FROM")
    smtp_host = os.environ.get("ACRIS_SMTP_HOST")
    smtp_port = int(os.environ.get("ACRIS_SMTP_PORT", "587"))
    smtp_user = os.environ.get("ACRIS_SMTP_USER")
    smtp_pass = os.environ.get("ACRIS_SMTP_PASS")

    if not (from_addr and smtp_host and smtp_user and smtp_pass):
        print("SMTP env vars not set — skipping email. Files were still "
              "written. See the header of this script for the variable names.")
        return

    by_borough = {}
    for r in new_rows:
        by_borough.setdefault(r["borough"], []).append(r)

    lines = [f"{len(new_rows)} new HUD Partial Claim(s) found across NYC.", ""]
    for borough, rows in sorted(by_borough.items()):
        lines.append(f"-- {borough} ({len(rows)}) --")
        for r in sorted(rows, key=lambda x: x["recorded_date"], reverse=True):
            amt = f"${r['doc_amount']}" if r["doc_amount"] else "amount n/a"
            lines.append(f"  {r['address'] or '(address n/a)'} {r['zip']}  "
                         f"| Block {r['block']} Lot {r['lot']}  "
                         f"| recorded {r['recorded_date']}  | {amt}")
        lines.append("")

    msg = MIMEText("\n".join(lines))
    msg["Subject"] = f"HUD Partial Claim Monitor — {len(new_rows)} new (NYC)"
    msg["From"] = from_addr
    msg["To"] = to_addr

    with smtplib.SMTP(smtp_host, smtp_port) as server:
        server.starttls()
        server.login(smtp_user, smtp_pass)
        server.sendmail(from_addr, [to_addr], msg.as_string())
    print(f"Emailed {to_addr}.")


def main():
    print("HUD Partial Claim Refresh — NYC (5 boroughs)\n")

    print("Step 1/5: finding HUD as lender in ACRIS Real Property Parties...")
    all_ids = fetch_hud_partial_claim_document_ids()
    print(f"  {len(all_ids)} document_ids found.\n")

    print("Step 2/5: pulling document details (type, amount, recorded date)...")
    master_info = fetch_master_info(all_ids)
    print(f"  {len(master_info)} are mortgage-family documents.\n")

    print("Step 3/5: resolving addresses and zip codes...")
    legals_info = fetch_legals_info(list(master_info.keys()))

    print("Step 4/5: pulling borrower name + mailing address "
          "(for the absentee-owner flag)...")
    borrower_info = fetch_borrower_info(list(master_info.keys()))
    all_rows = build_rows(master_info, legals_info, borrower_info)
    absentee_count = sum(1 for r in all_rows if r["absentee_owner"])
    print(f"  {len(all_rows)} property rows built "
          f"({absentee_count} flagged absentee-owner).\n")

    print("Step 5/5: writing output files and checking what's new...")
    write_outputs(all_rows)
    print(f"  Wrote {JSON_OUTPUT_PATH.name} and {CSV_OUTPUT_PATH.name} "
          f"({len(all_rows)} rows).")

    seen = load_seen_ids()
    current_ids = {r["document_id"] for r in all_rows}
    new_rows = [r for r in all_rows if r["document_id"] not in seen]
    if new_rows:
        print(f"  {len(new_rows)} new since last run.")
        send_email(new_rows)
    else:
        print("  Nothing new since last run.")
    save_seen_ids(current_ids)


if __name__ == "__main__":
    try:
        main()
    except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as e:
        print("\nNYC Open Data's website didn't respond in time, even after "
              "retrying several times. This is usually just their server "
              "being slow, not a problem with your computer or this script.",
              file=sys.stderr)
        print("No files were changed. Just run the script again "
              "(python3 hud_partial_claim_refresh.py) — it will start over "
              "and try again.", file=sys.stderr)
        sys.exit(1)
    except requests.HTTPError as e:
        print(f"NYC Open Data request failed: {e}", file=sys.stderr)
        sys.exit(1)
