"""
devashish_TCO.py

AWS Billing PDF → 4 Services → GCP TCO Mapping
Each service gets its own Excel sheet with identical column format.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PDF STRUCTURE — STRUCTURAL RULES (never change, values vary)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Every AWS billing PDF follows this fixed 4-level hierarchy:

LEVEL 0 — SERVICE HEADER
  Rule: Short line. Contains ONLY the service name + USD total.
        NO usage-dimension code. NO quantity.
  Examples (values change, structure doesn't):
    "EC2 Container Registry (ECR)   USD 18.00"
    "Redshift                       USD 573.25"
    "Secrets Manager                USD 0.00"
    "Step Functions                 USD 0.00"

LEVEL 1 — REGION
  Rule: Starts with a known AWS region display name.
        Cost = region subtotal (we ignore it).
        Credits show as (USD x.xx) in parentheses → skip.
  Examples:
    "Asia Pacific (Mumbai)          USD 573.25"
    "EU (Stockholm)                 USD 0.00"
    "No Region                   (USD 1.20)"   ← skip (credit)

LEVEL 2 — USAGE DIMENSION  ← WE EXTRACT COST FROM HERE
  Rule: Full AWS service name + usage-dimension code + USD cost.
        The usage-dimension code is ALWAYS the last token before "USD"
        and ALWAYS contains at least one hyphen (-) or colon (:).
        Cost on this line is the AUTHORITATIVE cost for that dimension.
  Examples (usage code values vary, structure doesn't):
    "Amazon EC2 Container Registry (ECR) APS3-TimedStorage-ByteHrs   USD 18.00"
    "Amazon Redshift APS3-RMS:Serverless                              USD 27.48"
    "Amazon Redshift RunServerlessCompute:001                         USD 545.77"
    "AWS Secrets Manager APS3-AWSSecretsManager-Secrets               USD 1.20"
    "AWS Step Functions APS3-StateTransition                          USD 0.00"

LEVEL 3 — DETAIL / SUB-LINE  ← WE EXTRACT QTY+UNIT FROM HERE
  Rule: Sits directly below a LEVEL-2 line.
        Contains the quantity, unit, and repeats the same cost.
        May start with "$" (pricing description) OR plain English.
        We read qty+unit ONLY — we already have the cost from LEVEL 2.
  Examples (values vary):
    "$0.10 per GB-month of data storage         179.997 GB-Mo    USD 18.00"
    "Storage charges with Redshift managed storage  1,052.989 GB-Mo  USD 27.48"
    "$0.4275 per RPU-Hr for Redshift serverless  1,276.663 RPU-Hr    USD 545.77"
    "$0.40 per Secret                            3 Secrets           USD 1.20"
    "$0 for first 4,000 state transitions        912 StateTransitions USD 0.00"

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
THE 3 UNIVERSAL RULES THE PARSER IS BUILT ON:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

RULE 1 — HOW TO IDENTIFY A LEVEL-2 LINE:
  The token immediately before "USD <cost>" at the end of the line
  contains at least one hyphen (-) or colon (:).
  This is the USAGE-DIMENSION CODE (e.g. APS3-TimedStorage-ByteHrs).
  If the last token before USD has no hyphen/colon → it's LEVEL 0 or 1.

RULE 2 — HOW TO IDENTIFY A LEVEL-3 (DETAIL) LINE:
  Comes right after a LEVEL-2 line.
  Contains a NUMBER followed by a UNIT WORD (qty pattern).
  May or may not start with "$".
  We extract qty+unit using a broad regex that catches any number+word pair.

RULE 3 — COST AUTHORITY:
  LEVEL 0 cost = service total          → IGNORED
  LEVEL 1 cost = region subtotal        → IGNORED
  LEVEL 2 cost = dimension cost         → EXTRACTED (authoritative)
  LEVEL 3 cost = same as LEVEL 2 cost   → IGNORED (already have it)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Usage:
    python devashish_TCO.py
    Place any AWS billing PDF inside the  input/  folder.
Output:
    output/<pdf_name>_TCO_Map.xlsx   (4 data sheets + Methodology)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""

from __future__ import annotations

import os
import re
import sys
from pathlib import Path
from datetime import datetime

import pdfplumber
import pandas as pd

try:
    import xlsxwriter
    HAS_XLSXWRITER = True
except ImportError:
    HAS_XLSXWRITER = False
    print("⚠️  xlsxwriter not found. Run: pip install xlsxwriter")

# =============================================================================
# SECTION 1 — FOLDERS
# =============================================================================

INPUT_FOLDER  = "input"
OUTPUT_FOLDER = "output"

# =============================================================================
# SECTION 2 — REGION MAPS  (structural — won't change)
# =============================================================================

USAGE_PREFIX_TO_REGION: dict[str, str] = {
    "USE1": "us-east-1",       "USE2": "us-east-2",
    "USW1": "us-west-1",       "USW2": "us-west-2",
    "APS1": "ap-southeast-1",  "APS2": "ap-southeast-2",
    "APS3": "ap-south-1",      "APS4": "ap-southeast-3",  "APS5": "ap-south-2",
    "APN1": "ap-northeast-1",  "APN2": "ap-northeast-2",  "APN3": "ap-northeast-3",
    "EUC1": "eu-central-1",    "EUC2": "eu-central-2",
    "EU":   "eu-west-1",       "EUW2": "eu-west-2",       "EUW3": "eu-west-3",
    "EUN1": "eu-north-1",      "EUS1": "eu-south-1",      "EUS2": "eu-south-2",
    "CAN1": "ca-central-1",    "SAE1": "sa-east-1",
    "MES1": "me-south-1",      "MEC1": "me-central-1",
}

LOCATION_TO_REGION: dict[str, str] = {
    "us east (n. virginia)":     "us-east-1",
    "us east (ohio)":            "us-east-2",
    "us west (n. california)":   "us-west-1",
    "us west (oregon)":          "us-west-2",
    "asia pacific (mumbai)":     "ap-south-1",
    "asia pacific (hyderabad)":  "ap-south-2",
    "asia pacific (singapore)":  "ap-southeast-1",
    "asia pacific (sydney)":     "ap-southeast-2",
    "asia pacific (tokyo)":      "ap-northeast-1",
    "asia pacific (seoul)":      "ap-northeast-2",
    "asia pacific (osaka)":      "ap-northeast-3",
    "eu (ireland)":              "eu-west-1",
    "eu (london)":               "eu-west-2",
    "eu (paris)":                "eu-west-3",
    "eu (frankfurt)":            "eu-central-1",
    "eu (stockholm)":            "eu-north-1",
    "canada (central)":          "ca-central-1",
    "south america (sao paulo)": "sa-east-1",
    "middle east (bahrain)":     "me-south-1",
    "middle east (uae)":         "me-central-1",
    "any":                        "Global",
    "global":                     "Global",
    "no region":                  "Global",
}

AWS_TO_GCP_REGION: dict[str, str] = {
    "us-east-1":      "us-east1",       "us-east-2":      "us-east4",
    "us-west-1":      "us-west2",       "us-west-2":      "us-west1",
    "ca-central-1":   "northamerica-northeast1",
    "eu-west-1":      "europe-west1",   "eu-west-2":      "europe-west2",
    "eu-west-3":      "europe-west9",   "eu-central-1":   "europe-west3",
    "eu-north-1":     "europe-north1",
    "ap-south-1":     "asia-south1",    "ap-south-2":     "asia-south2",
    "ap-southeast-1": "asia-southeast1","ap-southeast-2": "australia-southeast1",
    "ap-northeast-1": "asia-northeast1","ap-northeast-2": "asia-northeast3",
    "ap-northeast-3": "asia-northeast2",
    "sa-east-1":      "southamerica-east1",
    "me-central-1":   "me-central1",    "me-south-1":     "me-west1",
    "Global":         "Global",
}

# =============================================================================
# SECTION 3 — SERVICE DEFINITIONS
# =============================================================================

SERVICES: dict[str, dict] = {
    "ECR": {
        # These substrings identify the service at LEVEL 0 and LEVEL 2
        # They are fixed AWS service names — won't change across bills
        "svc_names":   ["ec2 container registry", "elastic container registry", "amazon ecr"],
        "aws_service": "Amazon ECR",
        "gcp_service": "Artifact Registry",
        "sheet_name":  "1. ECR",
        "rate_label":  "$0.10 / GB-Mo",
    },
    "Redshift": {
        "svc_names":   ["amazon redshift", "redshift"],
        "aws_service": "Amazon Redshift",
        "gcp_service": "BigQuery",
        "sheet_name":  "2. Redshift",
        "rate_label":  "Storage $0.02/GB-Mo | Compute $0.04/RPU-Hr",
    },
    "Secrets Manager": {
        "svc_names":   ["secrets manager", "aws secrets manager"],
        "aws_service": "AWS Secrets Manager",
        "gcp_service": "Secret Manager",
        "sheet_name":  "3. Secret Manager",
        "rate_label":  "Secrets $0.06/secret | API $0.03/10K ops",
    },
    "Step Functions": {
        "svc_names":   ["step functions", "aws step functions"],
        "aws_service": "AWS Step Functions",
        "gcp_service": "Cloud Workflows",
        "sheet_name":  "4. Step Functions",
        "rate_label":  "$0.01 / 1K steps (first 5K free/mo)",
    },
    "WAF": {
        "svc_names":   ["aws waf", "waf"],
        "aws_service": "AWS WAF",
        "gcp_service": "Cloud Armor",
        "sheet_name":  "5. WAF",
        "rate_label":  "Policy $5/mo | Rule $1/mo | Req $0.75/M",
    },
}

# All service names combined — used to detect when one section ends
ALL_SVC_NAMES: list[str] = [n for s in SERVICES.values() for n in s["svc_names"]]

# Non-target AWS services — seeing these also ends current section
OTHER_AWS_SERVICES: list[str] = [
    "elastic compute cloud", "amazon s3", "amazon rds", "amazon cloudfront",
    "amazon route 53", "amazon dynamodb", "aws lambda", "amazon sns",
    "amazon sqs", "amazon elastic container service", "amazon eks",
    "aws data transfer", "aws support", "amazon vpc", "amazon guardduty",
    "aws config", "amazon cloudwatch",
]

CUD_1YR = 0.80   # pay 80% of PAYG (20% off)
CUD_3YR = 0.60   # pay 60% of PAYG (40% off)

# =============================================================================
# SECTION 4 — STRUCTURAL REGEX  (built on rules, not specific values)
# =============================================================================

# ── Regex 1: Extract USD cost from end of any line ────────────────────────────
# Handles:  "... USD 18.00"         → positive cost
#           "... (USD 1.20)"        → credit (negative, skip)
COST_AT_END_RE = re.compile(
    r"\(USD\s+([\d,]+\.?\d*)\)\s*$"      # credit form
    r"|"
    r"USD\s+([\d,]+\.?\d*)\s*$",          # normal form
    re.IGNORECASE
)

# ── Regex 2: Detect region header lines (LEVEL 1) ────────────────────────────
# Fixed AWS region display name prefixes — these never change
REGION_LINE_RE = re.compile(
    r"^(Asia Pacific|US East|US West|EU|Canada|South America|Middle East|Africa|No Region|Any|Global)"
    r"(\s*\([^)]*\))?",
    re.IGNORECASE
)

# ── Regex 3: RULE 1 — Identify LEVEL-2 lines ────────────────────────────────
# The usage-dimension code is the last token before "USD x.xx"
# It ALWAYS contains at least one hyphen or colon.
# We capture it and the cost together.
#
# Pattern:  <anything>  <CODE_WITH_HYPHEN_OR_COLON>  USD  <cost>
# The CODE token: starts with a letter/digit, contains - or :, ends before USD
LEVEL2_LINE_RE = re.compile(
    r"(.+?)\s+"                           # service name prefix (greedy, up to code)
    r"([A-Za-z0-9][A-Za-z0-9_\.]*"       # code: starts with alphanumeric
    r"(?:[-:][A-Za-z0-9][A-Za-z0-9_\.\-:]*)+)"  # must have at least one -or: segment
    r"\s+USD\s+([\d,]+\.?\d*)\s*$",
    re.IGNORECASE
)

# ── Regex 4: Extract qty + unit from LEVEL-3 detail lines ───────────────────
# Broad pattern: catches ANY number followed by ANY word(s) as the unit.
# We use a known-unit list for precision but fall back broadly.
# "1,052.989 GB-Mo"  "1,276.663 RPU-Hr"  "3 Secrets"  "912 StateTransitions"
# "179.997 GB-Mo"    "2 API Requests"     "750,000 StateTransitions"
QTY_UNIT_RE = re.compile(
    r"([\d,]+\.?\d*)\s+"                  # quantity (with optional commas)
    r"([A-Za-z][A-Za-z0-9\-]*"           # unit: starts with letter
    r"(?:\s+[A-Za-z][A-Za-z]*)?)",        # optional second word (e.g. "API Requests")
    re.IGNORECASE
)

# Known units to prefer when multiple qty matches exist on same line
KNOWN_UNITS = {
    "gb-mo", "gb-month", "rpu-hr", "rpu-hour", "hrs", "hours",
    "secrets", "secret", "api requests", "api request",
    "requests", "request", "statetransitions", "statetransition",
    "transitions", "transition", "steps", "step", "count",
    "nodehrs", "vcpu-hours", "units",
}

# =============================================================================
# SECTION 5 — PARSER HELPER FUNCTIONS
# =============================================================================

def _extract_cost(line: str) -> float | None:
    """Extract USD cost from end of line. Returns negative for credits."""
    m = COST_AT_END_RE.search(line)
    if not m:
        return None
    if m.group(1) is not None:       # credit: (USD x.xx)
        return -float(m.group(1).replace(",", ""))
    return float(m.group(2).replace(",", ""))


def _is_region_line(line: str) -> bool:
    return bool(REGION_LINE_RE.match(line.strip()))


def _get_region_display(line: str) -> str:
    m = REGION_LINE_RE.match(line.strip())
    return m.group(0).strip() if m else line.strip()


def _which_service(lower_line: str) -> str | None:
    """Return service key if line contains that service's name."""
    for svc_key, svc_def in SERVICES.items():
        for name in svc_def["svc_names"]:
            if name in lower_line:
                return svc_key
    return None


# Short unit tokens that contain hyphens/colons but are NOT billing codes
# These would otherwise fool the LEVEL2_LINE_RE
SHORT_UNITS = {
    "gb-mo", "gb-month", "rpu-hr", "rpu-hour", "gb-s",
    "cpu-hr", "io-req", "api-req",
}

def _is_level2_line(line: str) -> tuple[str, str, float] | None:
    """
    RULE 1: Detect a LEVEL-2 usage-dimension line.

    Structural rule: the token immediately before "USD <cost>" at line end
    contains at least one hyphen (-) or colon (:) AND is at least 7 characters.

    Short hyphenated tokens like "GB-Mo" (5 chars) or "RPU-Hr" (6 chars)
    are units, not billing dimension codes — excluded by length check.

    Returns (service_prefix, usage_code, cost) if matched, else None.
    """
    m = LEVEL2_LINE_RE.match(line.strip())
    if not m:
        return None

    prefix     = m.group(1).strip()
    usage_code = m.group(2).strip()
    cost       = float(m.group(3).replace(",", ""))

    # Reject known short unit abbreviations (GB-Mo, RPU-Hr, etc.)
    if usage_code.lower() in SHORT_UNITS:
        return None

    # Reject any code shorter than 7 characters — real billing codes are always longer
    # e.g. "GB-Mo" = 5 chars, "RPU-Hr" = 6 chars, shortest real code = "EU-Node" = 7 chars
    if len(usage_code) < 7:
        return None

    return prefix, usage_code, cost


def _extract_qty_unit(line: str) -> tuple[float, str] | None:
    """
    RULE 2: Extract quantity and unit from a LEVEL-3 detail line.

    Finds ALL number+unit pairs on the line and returns the best match.
    "Best" = matches a known unit, or the largest number (most likely the qty).

    Works for any future unit types because it matches broadly then filters.
    The unit is cleaned — trailing "USD" or cost artifacts are stripped.
    """
    # Strip the trailing cost ("USD x.xx") first so it doesn't bleed into unit names
    clean_line = COST_AT_END_RE.sub("", line).strip()

    matches = QTY_UNIT_RE.findall(clean_line)
    if not matches:
        return None

    def clean_unit(u: str) -> str:
        """Remove trailing USD or numbers that bleed into unit strings."""
        u = re.split(r"\s+USD\b|\s+\d", u, flags=re.IGNORECASE)[0]
        return u.strip()

    # Step 1: Build exclusion sets FIRST before any selection
    # Pricing denominators: numbers that follow "per" e.g. "per 10000" in "$0.05 per 10000 API Requests"
    per_numbers = set(re.findall(r"\bper\s+(\d[\d,]*)", clean_line, re.IGNORECASE))
    # Noise unit words that are not real units
    NOISE_UNITS = {"per", "for", "in", "of", "at", "the", "a", "an"}

    # Step 2: Filter to real quantity matches only
    real_matches = [(q, u) for q, u in matches
                    if q not in per_numbers
                    and clean_unit(u).lower() not in NOISE_UNITS]

    candidate_matches = real_matches if real_matches else matches

    # Step 3: Among candidates, prefer known units (most reliable)
    for qty_str, unit in candidate_matches:
        unit_clean = clean_unit(unit)
        if unit_clean.lower() in KNOWN_UNITS:
            return float(qty_str.replace(",", "")), unit_clean

    # Step 4: Fallback — return candidate with smallest qty
    # (real usage counts are usually smaller than pricing denominators)
    if candidate_matches:
        best_qty, best_unit = min(candidate_matches, key=lambda x: float(x[0].replace(",", "")))
        return float(best_qty.replace(",", "")), clean_unit(best_unit)

    return None


def _resolve_region(usage_code: str, display_region: str) -> tuple[str, str]:
    """
    Resolve (aws_region_code, gcp_region) from usage code prefix or display name.

    Strategy 1: Extract prefix from usage code (e.g. "APS3" from "APS3-RMS:Serverless")
    Strategy 2: Match display name (e.g. "Asia Pacific (Mumbai)")
    """
    # Strategy 1: prefix before first hyphen or colon
    m = re.match(r"^([A-Z][A-Z0-9]+)[-:]", usage_code, re.IGNORECASE)
    if m:
        prefix = m.group(1).upper()
        if prefix in USAGE_PREFIX_TO_REGION:
            aws = USAGE_PREFIX_TO_REGION[prefix]
            return aws, AWS_TO_GCP_REGION.get(aws, f"unmapped:{aws}")

    # Strategy 2: display name
    aws = LOCATION_TO_REGION.get(display_region.lower().strip())
    if aws:
        return aws, AWS_TO_GCP_REGION.get(aws, f"unmapped:{aws}")

    return "unknown", "unknown"

# =============================================================================
# SECTION 6 — STATE-MACHINE PARSER
# =============================================================================

def extract_all_services(pdf_path: str) -> dict[str, list[dict]]:
    """
    Single-pass state-machine parser.

    State:
        current_svc     — which service section we're inside (or None)
        current_region  — current region display name
        after_level2    — True immediately after storing a LEVEL-2 record
                          (next line may be the LEVEL-3 detail)
        last_key        — (svc_key, list_index) of last stored record

    Logic per line (in order of precedence):
        1. Skip blank lines and credit lines
        2. If after_level2 → try to extract qty+unit (LEVEL-3), then clear flag
        3. Detect LEVEL-0 service header (has service name but NO usage code)
        4. Detect LEVEL-1 region line
        5. Detect section boundary (different service starts)
        6. Detect LEVEL-2 usage-dimension line (has usage code before USD)
           → extract usage_code + cost, set after_level2 = True
    """
    print(f"\n🧠 Parsing PDF: {pdf_path}")
    records: dict[str, list[dict]] = {svc: [] for svc in SERVICES}

    current_svc    = None
    current_region = "Global"
    after_level2   = False
    last_key       = None

    with pdfplumber.open(pdf_path) as pdf:
        print(f"   Pages: {len(pdf.pages)}")

        for page_num, page in enumerate(pdf.pages, 1):
            text = page.extract_text()
            if not text:
                continue

            for raw_line in text.split("\n"):
                line       = raw_line.strip()
                lower_line = line.lower()

                if not line:
                    continue

                # ── 1. Skip credits and activation lines ─────────────────────
                # "(USD x.xx)" = credit  |  "Credit" label  |  "AWS Activate"
                if re.search(r"\(usd\s+[\d,]+", lower_line):
                    after_level2 = False
                    continue
                if re.search(r"\bcredit\b|\baws activate\b", lower_line):
                    after_level2 = False
                    continue

                # ── 2. LEVEL-3 detail line ────────────────────────────────────
                # We just stored a LEVEL-2 record — next line has qty+unit.
                # This is the ONLY line where qty+unit lives.
                # Key insight: LEVEL-3 can start with "$" OR plain English
                # (e.g. "Storage charges with Redshift managed storage 1,052.989 GB-Mo")
                # We don't care what it starts with — we just look for the qty pattern.
                if after_level2 and last_key:
                    result = _extract_qty_unit(line)
                    if result:
                        qty, unit = result
                        svc_key, idx = last_key
                        records[svc_key][idx]["qty"]  = qty
                        records[svc_key][idx]["unit"] = unit
                    # Regardless of whether we found qty, this line is consumed.
                    # The detail line is always exactly ONE line.
                    after_level2 = False
                    continue

                # ── 3. LEVEL-0 service header detection ───────────────────────
                # A LEVEL-0 line has the service name but NO usage-dimension code.
                # We detect this by: service name present AND no hyphen/colon token
                # before "USD" (which would make it a LEVEL-2 line instead).
                #
                # We test for LEVEL-2 FIRST — if it IS a LEVEL-2 line, it's not a header.
                level2_result = _is_level2_line(line)

                svc_in_line = _which_service(lower_line)

                if svc_in_line and not level2_result:
                    # LEVEL-0: service name present, no usage code → it's a header
                    if current_svc != svc_in_line:
                        current_svc    = svc_in_line
                        current_region = "Global"
                        after_level2   = False
                    continue   # LEVEL-0 cost is the service total → ignore

                # ── 4. LEVEL-1 region line ────────────────────────────────────
                if _is_region_line(line):
                    current_region = _get_region_display(line)
                    after_level2   = False
                    continue   # Region subtotal → ignore

                # ── 5. Section boundary: different service begins ─────────────
                # If we see another service's name without a usage code,
                # that's a new LEVEL-0 header → switch service context.
                if svc_in_line and svc_in_line != current_svc and not level2_result:
                    current_svc    = svc_in_line
                    current_region = "Global"
                    after_level2   = False
                    continue

                # Check non-target services (exit current section)
                if current_svc:
                    for other in OTHER_AWS_SERVICES:
                        if other in lower_line and not level2_result:
                            if not any(n in lower_line
                                       for n in SERVICES[current_svc]["svc_names"]):
                                current_svc  = None
                                after_level2 = False
                                break

                if not current_svc:
                    continue

                # ── 6. LEVEL-2 usage-dimension line ──────────────────────────
                # RULE 1: has a usage code (token with hyphen/colon) before USD.
                #
                # CRITICAL OWNERSHIP CHECK:
                # A LEVEL-2 line must contain the CURRENT service's own name.
                # e.g. "Amazon Redshift APS3-RMS:Serverless USD 27.48"  ← has "redshift"
                #      "APS3-Fargate-ARM-GB-Hours  ... USD 5.06"        ← NO "redshift" → reject
                #
                # This prevents lines from OTHER services (Fargate, ECS, Lambda etc.)
                # from being captured under the wrong service section when the section
                # boundary was missed.
                if level2_result:
                    _, usage_code, cost = level2_result

                    # Ownership check: line must contain this service's name
                    svc_names = SERVICES[current_svc]["svc_names"]
                    line_belongs_to_current = any(n in lower_line for n in svc_names)

                    if not line_belongs_to_current:
                        # This LEVEL-2 line belongs to a different service.
                        # Exit the current section — we've passed into another service's block.
                        current_svc  = None
                        after_level2 = False
                        continue

                    aws_region, gcp_region = _resolve_region(usage_code, current_region)

                    records[current_svc].append({
                        "region":     current_region,
                        "usage_code": usage_code,
                        "cost":       cost,
                        "qty":        0.0,   # filled by next LEVEL-3 line
                        "unit":       "",    # filled by next LEVEL-3 line
                        "aws_region": aws_region,
                        "gcp_region": gcp_region,
                    })
                    last_key     = (current_svc, len(records[current_svc]) - 1)
                    after_level2 = True   # next line = detail
                    continue

    # ── Summary ───────────────────────────────────────────────────────────────
    print(f"\n   Extraction complete:")
    for svc, recs in records.items():
        total = sum(r["cost"] for r in recs if r["cost"] > 0)
        print(f"   {'✔' if recs else '–'}  {svc:<20} {len(recs):>3} rows   "
              f"AWS Total: ${total:.4f}")
        for r in recs:
            qty_s = f"{r['qty']:.4f} {r['unit']}" if r["qty"] else "—"
            print(f"         {r['region']:<28}  {r['usage_code']:<42}  "
                  f"qty: {qty_s:<22}  cost: ${r['cost']:.4f}")

    return records

# =============================================================================
# SECTION 7 — GCP COST CALCULATORS
# =============================================================================

def _gcp_ecr(qty: float, cost: float, unit: str) -> tuple[float, float, float]:
    """
    ECR → Artifact Registry
    Both charge $0.10/GB-month — direct 1:1 mapping.
    Use actual GB qty from the detail line; fall back to cost/rate.
    """
    gb   = qty if "gb" in unit.lower() and qty > 0 else (cost / 0.10 if cost > 0 else 0)
    payg = round(gb * 0.10, 4)
    return payg, round(payg * CUD_1YR, 4), round(payg * CUD_3YR, 4)


def _gcp_redshift(qty: float, cost: float, unit: str) -> tuple[float, float, float]:
    """
    Redshift → BigQuery
    Two Serverless dimensions (identified by unit):
      GB-Mo  → BigQuery active storage  $0.02/GB-Mo
      RPU-Hr → BigQuery Enterprise compute  $0.04/slot-hr
    Provisioned (node-hours) fallback → on-demand $5/TB (estimate 10GB/node-hr)
    """
    unit_l = unit.lower()
    if "gb" in unit_l and qty > 0:
        payg = round(qty * 0.02, 4)
    elif "rpu" in unit_l and qty > 0:
        payg = round(qty * 0.04, 4)
    else:
        # Provisioned: estimate TB from node-hours
        hrs  = qty if qty > 0 else (cost / 0.64 if cost > 0 else 0)
        payg = round((hrs * 10 / 1024) * 5.0, 4)
    return payg, round(payg * CUD_1YR, 4), round(payg * CUD_3YR, 4)


def _gcp_secrets(qty: float, cost: float, unit: str) -> tuple[float, float, float]:
    """
    Secrets Manager → Secret Manager
    Secrets:   AWS $0.40/secret/mo  →  GCP $0.06/secret/mo
    API calls: AWS $0.05/10K        →  GCP $0.03/10K
    Uses qty+unit from detail line; back-calculates from cost if unit unknown.
    """
    unit_l = unit.lower()
    if "secret" in unit_l and qty > 0:
        payg = round(qty * 0.06, 4)
    elif ("request" in unit_l or "api" in unit_l) and qty > 0:
        payg = round((qty / 10_000) * 0.03, 4)
    else:
        # Back-calculate: assume it's secrets if we have cost
        payg = round((cost / 0.40) * 0.06, 4) if cost > 0 else 0
    return payg, round(payg * CUD_1YR, 4), round(payg * CUD_3YR, 4)


def _gcp_stepfn(qty: float, cost: float, unit: str) -> tuple[float, float, float]:
    """
    Step Functions → Cloud Workflows
    AWS Standard: $0.025/1K transitions
    GCP: $0.01/1K steps (first 5,000 steps/month free)
    """
    transitions = qty if qty > 0 else ((cost / 0.025) * 1000 if cost > 0 else 0)
    billable    = max(0.0, transitions - 5_000)
    payg        = round((billable / 1_000) * 0.01, 4)
    return payg, round(payg * CUD_1YR, 4), round(payg * CUD_3YR, 4)


def _gcp_waf(qty: float, cost: float, unit: str,
             usage_code: str) -> tuple[float, float, float]:
    """
    AWS WAF → Cloud Armor
      WebACLV2  → Security Policy  AWS $5/ACL-mo   GCP $5/policy-mo  (1:1)
      RuleV2    → Rule             AWS $1/rule-mo  GCP $1/rule-mo    (1:1)
      BotControl subscription      AWS $10/mo      GCP ~$5/mo (Adaptive Protection est.)
      AntiDDoS subscription        AWS $20/mo      GCP $0 (included in Armor Standard)
      *-Request / Tier1            AWS $0.60/M req GCP $0.75/M req
    """
    code_l = usage_code.lower()
    u      = unit.lower()

    if "webacl" in code_l:
        count = qty if qty > 0 else (cost / 5.0 if cost > 0 else 0)
        payg  = round(count * 5.0, 4)
    elif "rulev2" in code_l and "request" not in code_l:
        count = qty if qty > 0 else (cost / 1.0 if cost > 0 else 0)
        payg  = round(count * 1.0, 4)
    elif "botcontrol" in code_l and "request" not in code_l:
        count = qty if qty > 0 else 1
        payg  = round(count * 5.0, 4)
    elif "antiddos" in code_l and "request" not in code_l:
        payg  = 0.0            # included in Cloud Armor Standard
    elif "request" in code_l or "tier" in code_l or "request" in u or "req" in u:
        reqs  = qty if qty > 0 else ((cost / 0.60) * 1_000_000 if cost > 0 else 0)
        payg  = round((reqs / 1_000_000) * 0.75, 4)
    else:
        payg  = round(cost * (0.75 / 0.60), 4) if cost > 0 else 0.0

    return payg, 0.0, 0.0   # no CUD for WAF/Cloud Armor


def _get_gcp_costs(svc_key: str, qty: float, cost: float,
                   unit: str, usage_code: str = "") -> tuple[float, float, float]:
    if svc_key == "ECR":             return _gcp_ecr(qty, cost, unit)
    if svc_key == "Redshift":        return _gcp_redshift(qty, cost, unit)
    if svc_key == "Secrets Manager": return _gcp_secrets(qty, cost, unit)
    if svc_key == "Step Functions":  return _gcp_stepfn(qty, cost, unit)
    if svc_key == "WAF":             return _gcp_waf(qty, cost, unit, usage_code)
    return cost, cost * CUD_1YR, cost * CUD_3YR

# =============================================================================
# SECTION 8 — BUILD DATAFRAMES  (identical columns for every service)
# =============================================================================

AWS_COLS  = ["AWS Region", "AWS Region Code", "AWS Usage Code",
             "AWS Quantity", "AWS Unit", "AWS Cost ($)", "AWS Service"]
GCP_COLS  = ["GCP Service", "GCP Region", "GCP PAYG ($)", "GCP Rate"]
COMP_COLS = ["Est. Savings ($)", "Recommendation"]
SPACERS   = [" ", "  "]
ALL_COLS  = AWS_COLS + SPACERS + GCP_COLS + COMP_COLS


def build_dataframe(svc_key: str, records: list[dict]) -> pd.DataFrame:
    svc_def = SERVICES[svc_key]
    rows    = []

    for r in records:
        if r["cost"] < 0:    # skip credits
            continue

        qty, cost, unit = r["qty"], r["cost"], r["unit"]
        usage_code       = r["usage_code"]
        payg, _, _       = _get_gcp_costs(svc_key, qty, cost, unit, usage_code)
        savings          = round(cost - payg, 4)

        rows.append({
            "AWS Region":       r["region"],
            "AWS Region Code":  r["aws_region"],
            "AWS Usage Code":   usage_code,
            "AWS Quantity":     round(qty, 6) if qty else 0.0,
            "AWS Unit":         unit,
            "AWS Cost ($)":     round(cost, 4),
            "AWS Service":      svc_def["aws_service"],
            "GCP Service":      svc_def["gcp_service"],
            "GCP Region":       r["gcp_region"],
            "GCP PAYG ($)":     payg,
            "GCP Rate":         svc_def["rate_label"],
            "Est. Savings ($)": savings,
            "Recommendation":   "GCP ✅" if savings >= 0 else "AWS ✅",
        })

    cols = AWS_COLS + GCP_COLS + COMP_COLS
    return pd.DataFrame(rows, columns=cols) if rows else pd.DataFrame(columns=cols)

# =============================================================================
# SECTION 9 — EXCEL EXPORTER  (same format for all 4 sheets)
# =============================================================================

COL_WIDTHS = [
    28, 16, 45, 14, 18, 14, 22,
    3,  3,
    22, 22, 14, 22,
    16, 14,
]
COL_FMTS = [
    None, None, None, "#,##0.000000", None, "$#,##0.0000", None,
    None, None,
    None, None, "$#,##0.0000", None,
    "$#,##0.0000", None,
]


def _write_sheet(workbook, svc_key: str, df: pd.DataFrame):
    svc_def = SERVICES[svc_key]

    aws_hdr   = workbook.add_format({"bold": True, "bg_color": "#FF9900",
                                      "font_color": "white", "border": 1,
                                      "align": "center", "valign": "vcenter", "text_wrap": True})
    gcp_hdr   = workbook.add_format({"bold": True, "bg_color": "#4285F4",
                                      "font_color": "white", "border": 1,
                                      "align": "center", "valign": "vcenter", "text_wrap": True})
    comp_hdr  = workbook.add_format({"bold": True, "bg_color": "#34A853",
                                      "font_color": "white", "border": 1,
                                      "align": "center", "valign": "vcenter", "text_wrap": True})
    spacer    = workbook.add_format({"bg_color": "#FFFFFF", "border": 0})
    title_fmt = workbook.add_format({"bold": True, "font_size": 13,
                                      "bg_color": "#2C3E50", "font_color": "white",
                                      "border": 1, "align": "center", "valign": "vcenter"})
    total_fmt = workbook.add_format({"bold": True, "num_format": "$#,##0.0000",
                                      "bg_color": "#FFF2CC", "border": 1, "align": "center"})
    total_lbl = workbook.add_format({"bold": True, "bg_color": "#FFF2CC", "border": 1})
    no_data   = workbook.add_format({"italic": True, "font_color": "#888888", "font_size": 11})
    green_bg  = "#E8F5E9"
    red_bg    = "#FFEBEE"
    even_bg   = "#F5F5F5"
    odd_bg    = "#FFFFFF"

    ws = workbook.add_worksheet(svc_def["sheet_name"][:31])

    # Title banner
    ws.set_row(0, 26)
    ws.merge_range(0, 0, 0, len(ALL_COLS) - 1,
                   f"  {svc_def['sheet_name']}  |  "
                   f"{svc_def['aws_service']}  →  {svc_def['gcp_service']}  |  TCO Mapping",
                   title_fmt)

    # Column headers
    ws.set_row(1, 22)
    for ci, col in enumerate(ALL_COLS):
        if col in AWS_COLS:   ws.write(1, ci, col, aws_hdr)
        elif col in GCP_COLS: ws.write(1, ci, col, gcp_hdr)
        elif col in COMP_COLS:ws.write(1, ci, col, comp_hdr)
        else:                 ws.write(1, ci, "", spacer)

    # Column widths
    for ci, (w, fmt_str) in enumerate(zip(COL_WIDTHS, COL_FMTS)):
        cell_fmt = workbook.add_format({"num_format": fmt_str}) if fmt_str else None
        ws.set_column(ci, ci, w, cell_fmt)

    if df.empty:
        ws.merge_range(2, 0, 2, len(ALL_COLS) - 1,
                       f"No billable data found for {svc_def['aws_service']} in this PDF.",
                       no_data)
        ws.freeze_panes(2, 0)
        return

    # Data rows
    df_out = df.copy()
    df_out[" "]  = ""
    df_out["  "] = ""
    df_out = df_out[ALL_COLS]

    for ri, row_data in enumerate(df_out.itertuples(index=False), start=2):
        savings_val = row_data[ALL_COLS.index("Est. Savings ($)")]
        row_bg      = even_bg if ri % 2 == 0 else odd_bg

        for ci, val in enumerate(row_data):
            col_name = ALL_COLS[ci]
            if col_name in ("Est. Savings ($)", "Recommendation"):
                bg = green_bg if savings_val >= 0 else red_bg
            else:
                bg = row_bg
            fmt_str  = COL_FMTS[ci]
            cell_fmt = workbook.add_format({
                "num_format": fmt_str or "General",
                "bg_color": bg, "border": 1,
                "align": "center", "valign": "vcenter",
            })
            ws.write(ri, ci, val, cell_fmt)

    # Totals row
    last_data = len(df_out) + 1
    total_row = last_data + 1
    ws.write(total_row, 0, "TOTALS", total_lbl)

    for col_name in ["AWS Cost ($)", "GCP PAYG ($)", "Est. Savings ($)"]:
        ci  = ALL_COLS.index(col_name)
        ltr = chr(ord("A") + ci)
        ws.write_formula(total_row, ci,
                         f"=SUM({ltr}3:{ltr}{last_data + 1})", total_fmt)

    ws.freeze_panes(2, 0)
    print(f"   ✔  Sheet '{svc_def['sheet_name']}' — {len(df)} rows")


def _write_methodology(workbook):
    ws = workbook.add_worksheet("Methodology")
    ws.set_column("A:A", 120)

    tf = workbook.add_format({"bold": True, "font_size": 14, "bg_color": "#D3D3D3", "border": 1})
    h1 = workbook.add_format({"bold": True, "font_size": 12, "font_color": "#1155cc"})
    bd = workbook.add_format({"font_size": 11, "text_wrap": True})
    mf = workbook.add_format({"font_size": 11, "font_color": "#38761d", "italic": True})

    r = 0
    ws.write(r, 0, "  FinOps TCO Mapping — ECR | Redshift | Secrets Manager | Step Functions | WAF", tf); r += 2
    ws.write(r, 0, f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", bd); r += 2

    ws.write(r, 0, "PDF Structure (4-Level Hierarchy — fixed across all billing periods)", h1); r += 1
    ws.write(r, 0, "  LEVEL 0 — Service Header:  'EC2 Container Registry (ECR)   USD 18.00'", bd); r += 1
    ws.write(r, 0, "    → Contains service name only. Cost = service total. IGNORED.", bd); r += 1
    ws.write(r, 0, "  LEVEL 1 — Region:           'Asia Pacific (Mumbai)          USD 18.00'", bd); r += 1
    ws.write(r, 0, "    → Region subtotal. IGNORED. Credits '(USD x.xx)' also skipped.", bd); r += 1
    ws.write(r, 0, "  LEVEL 2 — Usage Dimension:  'Amazon Redshift APS3-RMS:Serverless  USD 27.48'", bd); r += 1
    ws.write(r, 0, "    → Service name + usage code + cost. Cost EXTRACTED here (authoritative).", bd); r += 1
    ws.write(r, 0, "  LEVEL 3 — Detail Sub-line:  'Storage charges ...  1,052.989 GB-Mo  USD 27.48'", bd); r += 1
    ws.write(r, 0, "    → Quantity + unit extracted. Cost repeated here but IGNORED (already have it).", bd); r += 2

    ws.write(r, 0, "Key Parsing Rule — How LEVEL 2 is detected (works for ANY billing period)", h1); r += 1
    ws.write(r, 0, "  The token immediately before 'USD <cost>' at line end contains '-' or ':'.", bd); r += 1
    ws.write(r, 0, "  This is always the usage-dimension code (e.g. APS3-RMS:Serverless).", bd); r += 1
    ws.write(r, 0, "  If the last token before USD has NO hyphen/colon → it is Level 0 or 1.", bd); r += 2

    ws.write(r, 0, "TABLE 1 — ECR → Artifact Registry", h1); r += 1
    ws.write(r, 0, "  Both charge $0.10/GB-month. Direct 1:1 mapping.", bd); r += 1
    ws.write(r, 0, "  • PAYG = GB-Mo × $0.10  |  1Yr CUD = PAYG × 0.80  |  3Yr CUD = PAYG × 0.60", mf); r += 2

    ws.write(r, 0, "TABLE 2 — Redshift Serverless → BigQuery", h1); r += 1
    ws.write(r, 0, "  Two billing dimensions detected by unit from the detail line:", bd); r += 1
    ws.write(r, 0, "  • GB-Mo  (RMS:Serverless storage)    → BigQuery active storage $0.02/GB-Mo", mf); r += 1
    ws.write(r, 0, "  • RPU-Hr (RunServerlessCompute)      → BigQuery Enterprise $0.04/slot-hr", mf); r += 1
    ws.write(r, 0, "  • Provisioned node-hours fallback    → on-demand $5.00/TB (est. 10GB/node-hr)", mf); r += 2

    ws.write(r, 0, "TABLE 3 — Secrets Manager → Secret Manager", h1); r += 1
    ws.write(r, 0, "  Detected by unit from detail line:", bd); r += 1
    ws.write(r, 0, "  • Secrets      → GCP $0.06/secret/month  (vs AWS $0.40)", mf); r += 1
    ws.write(r, 0, "  • API Requests → GCP $0.03/10K ops        (vs AWS $0.05/10K)", mf); r += 2

    ws.write(r, 0, "TABLE 4 — Step Functions → Cloud Workflows", h1); r += 1
    ws.write(r, 0, "  AWS $0.025/1K state transitions  →  GCP $0.01/1K steps (first 5K free/month)", bd); r += 1
    ws.write(r, 0, "  • Billable = max(0, transitions − 5,000)", mf); r += 1
    ws.write(r, 0, "  • PAYG = (billable / 1,000) × $0.01", mf); r += 2

    ws.write(r, 0, "TABLE 5 — AWS WAF → Cloud Armor", h1); r += 1
    ws.write(r, 0, "  Web ACL:        AWS $5.00/ACL-mo       →  GCP $5.00/security-policy-mo  (1:1)", mf); r += 1
    ws.write(r, 0, "  Rule:           AWS $1.00/rule-mo       →  GCP $1.00/rule-mo             (1:1)", mf); r += 1
    ws.write(r, 0, "  Requests (Std): AWS $0.60/M req         →  GCP $0.75/M req", mf); r += 1
    ws.write(r, 0, "  Bot Control:    AWS $10/mo subscription  →  GCP ~$5/mo (Adaptive Protection est.)", mf); r += 1
    ws.write(r, 0, "  Anti-DDoS:      AWS $20/mo subscription  →  GCP $0.00 (included in Cloud Armor Standard)", mf); r += 1
    ws.write(r, 0, "  Bot/DDoS reqs:  AWS $0.15-$1.00/M req   →  GCP $0.75/M req", mf); r += 2
    ws.write(r, 0, "NOTE: 1-Year CUD and 3-Year CUD are NOT shown.", h1); r += 1
    ws.write(r, 0, "  CUDs apply only to compute (VMs/TPUs). These 5 services are usage-billed with no CUD pricing on GCP.", bd); r += 1


def export_to_excel(all_dfs: dict[str, pd.DataFrame], output_path: str):
    if not HAS_XLSXWRITER:
        for svc_key, df in all_dfs.items():
            df.to_csv(output_path.replace(".xlsx", f"_{svc_key}.csv"), index=False)
        return

    print(f"\n📁 Building Excel workbook: {output_path}")
    try:
        writer   = pd.ExcelWriter(output_path, engine="xlsxwriter")
        workbook = writer.book
        for svc_key in SERVICES:
            _write_sheet(workbook, svc_key, all_dfs[svc_key])
        _write_methodology(workbook)
        print("   ✔  Sheet 'Methodology'")
        workbook.close()
        print(f"\n✅ Excel saved → {output_path}")
    except PermissionError:
        print(f"\n❌ '{output_path}' is open. Close Excel and run again.")
        sys.exit(1)

# =============================================================================
# SECTION 10 — DEMO DATA  (exact values from your screenshots)
# =============================================================================

DEMO_RECORDS: dict[str, list[dict]] = {
    "ECR": [
        {"region": "Asia Pacific (Mumbai)", "usage_code": "APS3-TimedStorage-ByteHrs",
         "cost": 18.00, "qty": 179.997, "unit": "GB-Mo",
         "aws_region": "ap-south-1", "gcp_region": "asia-south1"},
        {"region": "US East (N. Virginia)", "usage_code": "TimedStorage-ByteHrs",
         "cost": 0.05,  "qty": 0.496,   "unit": "GB-Mo",
         "aws_region": "us-east-1",  "gcp_region": "us-east1"},
    ],
    "Redshift": [
        {"region": "Asia Pacific (Mumbai)", "usage_code": "APS3-RMS:Serverless",
         "cost": 27.48,  "qty": 1052.989, "unit": "GB-Mo",
         "aws_region": "ap-south-1", "gcp_region": "asia-south1"},
        {"region": "Asia Pacific (Mumbai)", "usage_code": "APS3-RunServerlessCompute:001",
         "cost": 545.77, "qty": 1276.663, "unit": "RPU-Hr",
         "aws_region": "ap-south-1", "gcp_region": "asia-south1"},
    ],
    "Secrets Manager": [
        {"region": "Asia Pacific (Mumbai)", "usage_code": "APS3-AWSSecretsManager-Secrets",
         "cost": 1.20, "qty": 3.0, "unit": "Secrets",
         "aws_region": "ap-south-1", "gcp_region": "asia-south1"},
        {"region": "Asia Pacific (Mumbai)", "usage_code": "APS3-AWSSecretsManagerAPIRequest",
         "cost": 0.00, "qty": 2.0, "unit": "API Requests",
         "aws_region": "ap-south-1", "gcp_region": "asia-south1"},
        {"region": "EU (Stockholm)",        "usage_code": "EUN1-AWSSecretsManagerAPIRequest",
         "cost": 0.00, "qty": 3.0, "unit": "API Requests",
         "aws_region": "eu-north-1", "gcp_region": "europe-north1"},
    ],
    "Step Functions": [
        {"region": "Asia Pacific (Mumbai)", "usage_code": "APS3-StateTransition",
         "cost": 0.00, "qty": 912.0, "unit": "StateTransitions",
         "aws_region": "ap-south-1", "gcp_region": "asia-south1"},
    ],
    "WAF": [
        {"region": "Any",                   "usage_code": "Global-WebACLV2",
         "cost": 20.00, "qty": 4.0,  "unit": "Month",
         "aws_region": "Global",     "gcp_region": "Global"},
        {"region": "Any",                   "usage_code": "Global-RuleV2",
         "cost": 17.00, "qty": 17.0, "unit": "Month",
         "aws_region": "Global",     "gcp_region": "Global"},
        {"region": "Any",                   "usage_code": "Global-RequestV2-Tier1",
         "cost": 12.05, "qty": 20082221.0, "unit": "Request",
         "aws_region": "Global",     "gcp_region": "Global"},
        {"region": "Asia Pacific (Mumbai)", "usage_code": "APS3-WebACLV2",
         "cost": 15.00, "qty": 3.0,  "unit": "Month",
         "aws_region": "ap-south-1", "gcp_region": "asia-south1"},
        {"region": "Asia Pacific (Mumbai)", "usage_code": "APS3-RuleV2",
         "cost":  9.00, "qty": 9.0,  "unit": "Month",
         "aws_region": "ap-south-1", "gcp_region": "asia-south1"},
        {"region": "Asia Pacific (Mumbai)", "usage_code": "APS3-RequestV2-Tier1",
         "cost": 54.26, "qty": 90437417.0, "unit": "Request",
         "aws_region": "ap-south-1", "gcp_region": "asia-south1"},
    ],
}

# =============================================================================
# SECTION 11 — MAIN
# =============================================================================

def main():
    print("\n" + "=" * 70)
    print("  AWS → GCP TCO  (5 Services, 7 Sheets)")
    print("  ECR | Redshift | Secrets Manager | Step Functions | WAF")
    print("  (No CUD — usage-billed services, CUDs not applicable)")
    print("=" * 70)

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    os.makedirs(INPUT_FOLDER,  exist_ok=True)

    pdf_path  = None
    demo_mode = False
    for f in sorted(os.listdir(INPUT_FOLDER)):
        if f.lower().endswith(".pdf"):
            pdf_path = os.path.join(INPUT_FOLDER, f)
            break

    if not pdf_path:
        print(f"\n⚠️  No PDF in '{INPUT_FOLDER}/' — DEMO MODE\n")
        demo_mode = True

    output_stem = Path(pdf_path).stem if pdf_path else "demo"
    output_path = os.path.join(OUTPUT_FOLDER, f"{output_stem}_TCO_Map.xlsx")
    print(f"\n{'[DEMO]' if demo_mode else f'[LIVE] PDF: {pdf_path}'}")

    print(f"\n--- Extracting Billing Data ---")
    raw = DEMO_RECORDS if demo_mode else extract_all_services(pdf_path)

    print(f"\n--- Calculating GCP Costs ---")
    all_dfs: dict[str, pd.DataFrame] = {}
    for svc_key in SERVICES:
        df = build_dataframe(svc_key, raw.get(svc_key, []))
        all_dfs[svc_key] = df
        if not df.empty:
            a, g = df["AWS Cost ($)"].sum(), df["GCP PAYG ($)"].sum()
            print(f"   {svc_key:<20}  {len(df):>2} rows   "
                  f"AWS ${a:>8.2f}  →  GCP ${g:>8.2f}  ({'GCP ✅' if g <= a else 'AWS ✅'})")
        else:
            print(f"   {svc_key:<20}   0 rows  (no data found)")

    print(f"\n--- Exporting to Excel ---")
    export_to_excel(all_dfs, output_path)

    g_aws  = sum(df["AWS Cost ($)"].sum() for df in all_dfs.values() if not df.empty)
    g_payg = sum(df["GCP PAYG ($)"].sum() for df in all_dfs.values() if not df.empty)

    print("\n" + "=" * 70)
    print(f"  {'Service':<22} {'AWS Cost':>10}  {'GCP PAYG':>10}  {'Savings':>10}")
    print(f"  {'─' * 58}")
    for svc_key, df in all_dfs.items():
        a = df["AWS Cost ($)"].sum()  if not df.empty else 0
        g = df["GCP PAYG ($)"].sum()  if not df.empty else 0
        print(f"  {svc_key:<22} ${a:>9.2f}   ${g:>9.2f}   ${a-g:>9.2f}")
    print(f"  {'─' * 58}")
    print(f"  {'GRAND TOTAL':<22} ${g_aws:>9.2f}   ${g_payg:>9.2f}   ${g_aws-g_payg:>9.2f}")
    print(f"\n  Output: {output_path}")
    print("=" * 70)
    print("\n✅ Process Complete!")


if __name__ == "__main__":
    main()