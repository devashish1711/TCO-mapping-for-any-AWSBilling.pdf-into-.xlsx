"""
devashish_TCO.py  —  AWS Billing PDF → GCP TCO Mapping
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Services:
  1. Amazon EC2 Container Registry (ECR)  →  Artifact Registry
  2. Amazon Redshift                       →  BigQuery
  3. AWS Secrets Manager                   →  Secret Manager
  4. AWS Step Functions                    →  Cloud Workflows
  5. AWS WAF                               →  Cloud Armor

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
AUTOMATIC PDF HANDLING — TWO MODES, ZERO MANUAL SWITCHING
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

MODE A — Normal text-based PDF  (any standard AWS billing export)
  ┌──────────────────────────────────────────────────────────────────┐
  │  PDF STRUCTURE (4 levels — fixed across ALL billing periods)     │
  │                                                                  │
  │  L0  "EC2 Container Registry (ECR)    USD 18.00"                 │
  │       → Service total. IGNORED (= sum of regions).               │
  │                                                                  │
  │  L1  "Asia Pacific (Mumbai)           USD 18.00"                 │
  │       → Region subtotal. IGNORED (= sum of dimensions).          │
  │                                                                  │
  │  L2  "Amazon ECR APS3-TimedStorage-ByteHrs   USD 18.00"          │
  │       → COST EXTRACTED HERE (authoritative).                     │
  │         Detected by: last token before USD contains '-' or ':'   │
  │         AND token length >= 7 chars (filters short units)         │
  │         AND ownership check: line must contain service name       │
  │                                                                  │
  │  L3  "$0.10 per GB-month ... 179.997 GB-Mo  USD 18.00"           │
  │       → QTY + UNIT extracted here only. Cost ignored (= L2).     │
  │         Two formats handled:                                      │
  │         Format A: "... 179.997 GB-Mo USD x.xx"  (unit after qty) │
  │         Format B: "state transition 3,281.00 USD x.xx" (no unit) │
  └──────────────────────────────────────────────────────────────────┘

MODE B — Image-based or Redacted PDF  (scanned / costs blacked out)
  Detected automatically (0 extractable text characters).
  → Fill REDACTED_DATA (Section 5) with quantities read visually.
  → AWS costs are back-calculated using public list prices.
  → GCP PAYG computed from the same quantities.
  → Orange warning banner added to every Excel sheet.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Usage:
  Place PDF in ./input/  → output written to ./output/
  No PDF found           → DEMO mode with sample data
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
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
    print("xlsxwriter not found — run: pip install xlsxwriter")

# =============================================================================
# SECTION 1 — FOLDERS
# =============================================================================

INPUT_FOLDER  = "input"
OUTPUT_FOLDER = "output"

# =============================================================================
# SECTION 2 — REGION MAPS  (structural, won't change across billing periods)
# =============================================================================

# Maps billing-code PREFIX → AWS region code
# e.g. "APS3-TimedStorage-ByteHrs" → prefix "APS3" → "ap-south-1"
USAGE_PREFIX_TO_REGION: dict[str, str] = {
    "USE1": "us-east-1",       "USE2": "us-east-2",
    "USW1": "us-west-1",       "USW2": "us-west-2",
    "APS1": "ap-southeast-1",  "APS2": "ap-southeast-2",
    "APS3": "ap-south-1",      "APS4": "ap-southeast-3",  "APS5": "ap-south-2",
    "APN1": "ap-northeast-1",  "APN2": "ap-northeast-2",  "APN3": "ap-northeast-3",
    "APE1": "ap-east-1",
    "EUC1": "eu-central-1",    "EUC2": "eu-central-2",
    "EU":   "eu-west-1",       "EUW2": "eu-west-2",       "EUW3": "eu-west-3",
    "EUN1": "eu-north-1",      "EUS1": "eu-south-1",      "EUS2": "eu-south-2",
    "CAN1": "ca-central-1",    "SAE1": "sa-east-1",
    "MES1": "me-south-1",      "MEC1": "me-central-1",
    "AFS1": "af-south-1",
}

# Maps region display name (from PDF) → AWS region code
LOCATION_TO_REGION: dict[str, str] = {
    "us east (n. virginia)":          "us-east-1",
    "us east (northern virginia)":    "us-east-1",
    "us east (ohio)":                 "us-east-2",
    "us west (n. california)":        "us-west-1",
    "us west (northern california)":  "us-west-1",
    "us west (oregon)":               "us-west-2",
    "asia pacific (mumbai)":          "ap-south-1",
    "asia pacific (hyderabad)":       "ap-south-2",
    "asia pacific (singapore)":       "ap-southeast-1",
    "asia pacific (sydney)":          "ap-southeast-2",
    "asia pacific (tokyo)":           "ap-northeast-1",
    "asia pacific (seoul)":           "ap-northeast-2",
    "asia pacific (osaka)":           "ap-northeast-3",
    "asia pacific (hong kong)":       "ap-east-1",
    "asia pacific (jakarta)":         "ap-southeast-3",
    "asia pacific (melbourne)":       "ap-southeast-4",
    "eu (ireland)":                   "eu-west-1",
    "eu (london)":                    "eu-west-2",
    "eu (paris)":                     "eu-west-3",
    "eu (frankfurt)":                 "eu-central-1",
    "eu (stockholm)":                 "eu-north-1",
    "canada (central)":               "ca-central-1",
    "south america (sao paulo)":      "sa-east-1",
    "middle east (bahrain)":          "me-south-1",
    "middle east (uae)":              "me-central-1",
    "africa (cape town)":             "af-south-1",
    "any":                            "Global",
    "global":                         "Global",
    "no region":                      "Global",
}

AWS_TO_GCP_REGION: dict[str, str] = {
    "us-east-1":      "us-east1",            "us-east-2":      "us-east4",
    "us-west-1":      "us-west2",            "us-west-2":      "us-west1",
    "ca-central-1":   "northamerica-northeast1",
    "eu-west-1":      "europe-west1",        "eu-west-2":      "europe-west2",
    "eu-west-3":      "europe-west9",        "eu-central-1":   "europe-west3",
    "eu-north-1":     "europe-north1",
    "ap-south-1":     "asia-south1",         "ap-south-2":     "asia-south2",
    "ap-southeast-1": "asia-southeast1",     "ap-southeast-2": "australia-southeast1",
    "ap-northeast-1": "asia-northeast1",     "ap-northeast-2": "asia-northeast3",
    "ap-northeast-3": "asia-northeast2",     "ap-east-1":      "asia-east2",
    "ap-southeast-3": "asia-southeast2",     "ap-southeast-4": "australia-southeast2",
    "sa-east-1":      "southamerica-east1",
    "me-central-1":   "me-central1",         "me-south-1":     "me-west1",
    "af-south-1":     "africa-south1",
    "Global":         "Global",
}

# =============================================================================
# SECTION 3 — SERVICE DEFINITIONS
# =============================================================================

SERVICES: dict[str, dict] = {
    "ECR": {
        # Substrings that identify this service in L0 and L2 lines
        # (checked against lowercase line — add any aliases AWS uses)
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
        "rate_label":  "Storage $0.02/GB-Mo | Compute $0.04/slot-hr",
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
        "rate_label":  "$0.01/1K steps (first 5K free/mo)",
    },
    "WAF": {
        "svc_names":   ["aws waf", "waf"],
        "aws_service": "AWS WAF",
        "gcp_service": "Cloud Armor",
        "sheet_name":  "5. WAF",
        "rate_label":  "Policy $5/mo | Rule $1/mo | Req $0.75/M",
    },
}

ALL_SVC_NAMES: list[str] = [n for s in SERVICES.values() for n in s["svc_names"]]

# =============================================================================
# SECTION 4 — AWS PUBLIC RATES  (used only in image/redacted mode)
# Update if AWS changes pricing. Maps unit keyword → $ per unit.
# =============================================================================

AWS_PUBLIC_RATES: dict[str, dict] = {
    "ECR": {
        "gb-mo": 0.10, "gb-month": 0.10, "default": 0.10,
    },
    "Redshift": {
        "gb-mo": 0.024, "gb-month": 0.024,
        "rpu-hr": 0.4375, "rpu-hour": 0.4375,
        "default": 0.4375,
    },
    "Secrets Manager": {
        "secrets": 0.40, "secret": 0.40,
        "api requests": 0.05 / 10_000, "api request": 0.05 / 10_000,
        "requests":     0.05 / 10_000, "request":     0.05 / 10_000,
        "default": 0.40,
    },
    "Step Functions": {
        "statetransitions": 0.025 / 1_000, "statetransition": 0.025 / 1_000,
        "transitions":      0.025 / 1_000, "transition":      0.025 / 1_000,
        "state transition":  0.025 / 1_000,
        "default":          0.025 / 1_000,
    },
    "WAF": {
        "month": 5.00, "months": 5.00, "webacl": 5.00, "webacls": 5.00,
        "rules": 1.00, "rule":   1.00,
        "requests": 0.60 / 1_000_000, "request": 0.60 / 1_000_000,
        "default": 1.00,
    },
}


def back_calc_aws_cost(svc_key: str, qty: float, unit: str) -> float:
    """Estimate AWS cost = qty × public list rate. Used in image/redacted mode only."""
    rates = AWS_PUBLIC_RATES.get(svc_key, {})
    rate  = rates.get(unit.lower().strip(), rates.get("default", 0.0))
    return round(qty * rate, 4)

# =============================================================================
# SECTION 5 — REDACTED PDF DATA
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Fill this when your PDF is image-based or has its cost column redacted.
# Read the "Usage Quantity" column visually from the PDF, enter below.
#
# Fields:
#   "region"     : region display name as shown in PDF (e.g. "EU (Ireland)")
#   "usage_code" : billing-dimension code from the bold label in PDF
#                  (e.g. "EU-RMS:Serverless", "USE1-StateTransition")
#   "qty"        : the number from the Usage Quantity column
#   "unit"       : unit for that qty (e.g. "GB-Mo", "Secrets", "Request")
#   "cost"       : always set 0.0 — auto back-calculated by the script
#
# SAMPLE: Semantico Redacted Bill (CKSG2526-3259)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
REDACTED_DATA: dict[str, list[dict]] = {
    "ECR": [
        {"region": "EU (Frankfurt)",        "usage_code": "EUC1-TimedStorage-ByteHrs",
         "qty": 0.69,    "unit": "GB-Mo",  "cost": 0.0},
        {"region": "EU (Ireland)",          "usage_code": "EU-TimedStorage-ByteHrs",
         "qty": 0.66,    "unit": "GB-Mo",  "cost": 0.0},
        {"region": "US East (N. Virginia)", "usage_code": "TimedStorage-ByteHrs",
         "qty": 1810.24, "unit": "GB-Mo",  "cost": 0.0},
        {"region": "US West (Oregon)",      "usage_code": "USW2-TimedStorage-ByteHrs",
         "qty": 3.56,    "unit": "GB-Mo",  "cost": 0.0},
    ],
    "Redshift": [
        {"region": "EU (Ireland)",          "usage_code": "EU-RMS:Serverless",
         "qty": 0.81,    "unit": "GB-Mo",  "cost": 0.0},
    ],
    "Secrets Manager": [
        {"region": "US East (N. Virginia)", "usage_code": "USE1-AWSSecretsManager-Secrets",
         "qty": 18.0,    "unit": "Secrets","cost": 0.0},
    ],
    "Step Functions": [
        {"region": "US East (N. Virginia)", "usage_code": "USE1-StateTransition",
         "qty": 3281.0,  "unit": "StateTransitions","cost": 0.0},
    ],
    "WAF": [
        {"region": "Asia Pacific (Sydney)", "usage_code": "APS2-RuleV2",
         "qty": 7.0,     "unit": "rule",   "cost": 0.0},
        {"region": "Asia Pacific (Sydney)", "usage_code": "APS2-WebACLV2",
         "qty": 2.0,     "unit": "Month",  "cost": 0.0},
        {"region": "EU (Ireland)",          "usage_code": "EU-RequestV2-Tier0",
         "qty": 22261047.0, "unit": "Request","cost": 0.0},
        {"region": "EU (Ireland)",          "usage_code": "EU-RuleV2",
         "qty": 13.0,    "unit": "rule",   "cost": 0.0},
        {"region": "EU (Ireland)",          "usage_code": "EU-WebACLV2",
         "qty": 4.0,     "unit": "Month",  "cost": 0.0},
        {"region": "US East (N. Virginia)", "usage_code": "USE1-RequestV2-Tier1",
         "qty": 17365.0, "unit": "Request","cost": 0.0},
        {"region": "US East (N. Virginia)", "usage_code": "USE1-RuleV2",
         "qty": 4.0,     "unit": "rule",   "cost": 0.0},
        {"region": "US East (N. Virginia)", "usage_code": "USE1-WebACLV2",
         "qty": 1.0,     "unit": "Month",  "cost": 0.0},
    ],
}

# =============================================================================
# SECTION 6 — DEMO DATA  (used when no PDF found in input/)
# =============================================================================

DEMO_RECORDS: dict[str, list[dict]] = {
    "ECR": [
        {"region": "Asia Pacific (Mumbai)", "usage_code": "APS3-TimedStorage-ByteHrs",
         "cost": 18.00, "qty": 179.997, "unit": "GB-Mo",
         "aws_region": "ap-south-1", "gcp_region": "asia-south1"},
        {"region": "US East (N. Virginia)", "usage_code": "TimedStorage-ByteHrs",
         "cost": 0.05, "qty": 0.496, "unit": "GB-Mo",
         "aws_region": "us-east-1", "gcp_region": "us-east1"},
    ],
    "Redshift": [
        {"region": "Asia Pacific (Mumbai)", "usage_code": "APS3-RMS:Serverless",
         "cost": 27.48, "qty": 1052.989, "unit": "GB-Mo",
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
    ],
    "Step Functions": [
        {"region": "Asia Pacific (Mumbai)", "usage_code": "APS3-StateTransition",
         "cost": 0.00, "qty": 912.0, "unit": "StateTransitions",
         "aws_region": "ap-south-1", "gcp_region": "asia-south1"},
    ],
    "WAF": [
        {"region": "Any", "usage_code": "Global-WebACLV2",
         "cost": 20.00, "qty": 4.0, "unit": "Month",
         "aws_region": "Global", "gcp_region": "Global"},
        {"region": "Any", "usage_code": "Global-RuleV2",
         "cost": 17.00, "qty": 17.0, "unit": "rule",
         "aws_region": "Global", "gcp_region": "Global"},
        {"region": "Any", "usage_code": "Global-RequestV2-Tier1",
         "cost": 12.05, "qty": 20082221.0, "unit": "Request",
         "aws_region": "Global", "gcp_region": "Global"},
        {"region": "Asia Pacific (Mumbai)", "usage_code": "APS3-WebACLV2",
         "cost": 15.00, "qty": 3.0, "unit": "Month",
         "aws_region": "ap-south-1", "gcp_region": "asia-south1"},
        {"region": "Asia Pacific (Mumbai)", "usage_code": "APS3-RuleV2",
         "cost": 9.00, "qty": 9.0, "unit": "rule",
         "aws_region": "ap-south-1", "gcp_region": "asia-south1"},
        {"region": "Asia Pacific (Mumbai)", "usage_code": "APS3-RequestV2-Tier1",
         "cost": 54.26, "qty": 90437417.0, "unit": "Request",
         "aws_region": "ap-south-1", "gcp_region": "asia-south1"},
    ],
}

# =============================================================================
# SECTION 7 — STRUCTURAL REGEX  (built on rules, not values — works any period)
# =============================================================================

# Trailing USD cost — handles positive "USD x.xx" and credits "(USD x.xx)"
COST_AT_END_RE = re.compile(
    r"\(USD\s+([\d,]+\.?\d*)\)\s*$"           # credit form → negative
    r"|"
    r"USD\s+([\d,]+\.?\d*)\s*$",              # normal form
    re.IGNORECASE
)

# LEVEL 1 — region header detection
# Covers all known AWS region display name prefixes
REGION_LINE_RE = re.compile(
    r"^(Asia Pacific|US East|US West|EU|Canada|South America|"
    r"Middle East|Africa|No Region|Any|Global)"
    r"(\s*\([^)]*\))?",
    re.IGNORECASE
)

# LEVEL 2 — usage dimension line detection
# Rule: last token before "USD <cost>" contains '-' or ':' AND is >= 7 chars
# This is the structural invariant that holds across ALL AWS billing PDFs
LEVEL2_LINE_RE = re.compile(
    r"(.+?)\s+"
    r"([A-Za-z0-9][A-Za-z0-9_\.]*"
    r"(?:[-:][A-Za-z0-9][A-Za-z0-9_\.\-:]*)+)"
    r"\s+USD\s+([\d,]+\.?\d*)\s*$",
    re.IGNORECASE
)

# Short hyphenated tokens that look like billing codes but are actually units
SHORT_UNITS = {
    "gb-mo", "gb-month", "rpu-hr", "rpu-hour", "gb-s",
    "cpu-hr", "io-req", "api-req", "vcpu-hr", "node-hr",
}

# Recognised unit strings — preferred in qty/unit matching
KNOWN_UNITS = {
    "gb-mo", "gb-month", "rpu-hr", "rpu-hour", "hrs", "hours",
    "secrets", "secret", "api requests", "api request",
    "requests", "request", "statetransitions", "statetransition",
    "transitions", "transition", "steps", "step", "count",
    "nodehrs", "vcpu-hours", "units", "month", "months",
    "webacls", "webacl", "rules", "rule", "resource-month",
}

# =============================================================================
# SECTION 8 — PDF DETECTION
# =============================================================================

def detect_pdf_mode(pdf_path: str) -> str:
    """
    Returns 'text' if pdfplumber can extract characters, else 'image'.
    Checks first 5 pages only for speed.
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            total_chars = sum(len(p.chars) for p in pdf.pages[:5])
        return "text" if total_chars > 0 else "image"
    except Exception:
        return "text"

# =============================================================================
# SECTION 9 — PARSER HELPER FUNCTIONS
# =============================================================================

def _extract_cost(line: str) -> float | None:
    """Extract USD cost from end of line. Returns None if no cost found."""
    m = COST_AT_END_RE.search(line)
    if not m:
        return None
    # credit → negative; normal → positive
    return (-float(m.group(1).replace(",", ""))
            if m.group(1) is not None
            else float(m.group(2).replace(",", "")))


def _is_region_line(line: str) -> bool:
    """True if the line is a LEVEL-1 region header."""
    s = line.strip()
    if not REGION_LINE_RE.match(s):
        return False
    # Must also end with a USD cost (region subtotal)
    return bool(COST_AT_END_RE.search(s))


def _get_region_display(line: str) -> str:
    """Extract the region display name from a LEVEL-1 line."""
    m = REGION_LINE_RE.match(line.strip())
    return m.group(0).strip() if m else line.strip()


def _which_service(lower_line: str) -> str | None:
    """Return the service key if the line mentions that service's name."""
    for svc_key, svc_def in SERVICES.items():
        for name in svc_def["svc_names"]:
            if name in lower_line:
                return svc_key
    return None


def _is_level2_line(line: str) -> tuple[str, str, float] | None:
    """
    RULE 1: Detect a LEVEL-2 usage-dimension line.

    Structural invariant across ALL AWS billing PDFs:
      The token immediately before "USD <cost>" at line end:
        - Contains at least one hyphen (-) or colon (:)
        - Is at least 7 characters long
          (filters short units: GB-Mo=5, RPU-Hr=6)
        - Is not a known short unit abbreviation

    Returns (service_prefix, usage_code, cost) or None.
    """
    m = LEVEL2_LINE_RE.match(line.strip())
    if not m:
        return None
    prefix     = m.group(1).strip()
    usage_code = m.group(2).strip()
    cost       = float(m.group(3).replace(",", ""))
    if usage_code.lower() in SHORT_UNITS:
        return None
    if len(usage_code) < 7:
        return None
    return prefix, usage_code, cost


def _extract_qty_unit(line: str) -> tuple[float, str] | None:
    """
    RULE 2: Extract quantity + unit from a LEVEL-3 detail line.

    Handles two real-world formats found in AWS billing PDFs:

    Format A (standard — most bills):
      "$0.10 per GB-month ... 179.997 GB-Mo USD 18.00"
      → qty=179.997, unit="GB-Mo"
      The unit word explicitly follows the quantity number.

    Format B (plain-English — some CK/reseller bills):
      "state transition 3,281.00 USD 0.00"
      → qty=3281, unit inferred from description keywords
      The quantity is the last number before USD, no explicit unit after it.

    Strategy:
      1. Strip trailing USD cost.
      2. Find all number+word pairs.
      3. Prefer pairs where the word is a known unit string.
      4. Fall back to smallest-number heuristic (usage < pricing denominators).
      5. If still nothing: grab last number before USD, infer unit from
         description keyword matching.
    """
    clean = COST_AT_END_RE.sub("", line).strip()

    # Step A — find all number+word pairs
    QTY_UNIT_RE_LOCAL = re.compile(
        r"([\d,]+\.?\d*)\s+([A-Za-z][A-Za-z0-9\-]*(?:\s+[A-Za-z][A-Za-z]*)?)",
        re.IGNORECASE
    )
    matches = QTY_UNIT_RE_LOCAL.findall(clean)

    def clean_unit(u: str) -> str:
        return re.split(r"\s+USD\b|\s+\d", u, flags=re.IGNORECASE)[0].strip()

    per_numbers = set(re.findall(r"\bper\s+(\d[\d,]*)", clean, re.IGNORECASE))
    NOISE = {"per", "for", "in", "of", "at", "the", "a", "an"}

    real = [(q, u) for q, u in matches
            if q not in per_numbers
            and clean_unit(u).lower() not in NOISE]
    candidates = real if real else matches

    # Step B — prefer known unit strings (most reliable)
    for qty_str, unit in candidates:
        u = clean_unit(unit)
        if u.lower() in KNOWN_UNITS:
            return float(qty_str.replace(",", "")), u

    # Step C — fallback: smallest qty (real usage < pricing denominators)
    if candidates:
        bq, bu = min(candidates, key=lambda x: float(x[0].replace(",", "")))
        return float(bq.replace(",", "")), clean_unit(bu)

    # Step D — Format B: last number before end of cleaned line
    last_num = re.search(r"([\d,]+\.?\d*)\s*$", clean)
    if last_num:
        qty   = float(last_num.group(1).replace(",", ""))
        clean_l = clean.lower()
        # Infer unit from description keywords
        if any(k in clean_l for k in ["state transition", "statetransition"]):
            return qty, "StateTransitions"
        if "resource-month" in clean_l or "resource month" in clean_l:
            return qty, "resource-month"
        if "web acl" in clean_l or "webacl" in clean_l:
            return qty, "Month"
        if "rule" in clean_l:
            return qty, "rule"
        if "secret" in clean_l:
            return qty, "Secrets"
        if "request" in clean_l:
            return qty, "Request"
        if "gb-month" in clean_l or "gb-mo" in clean_l:
            return qty, "GB-Mo"
        if "rpu" in clean_l:
            return qty, "RPU-Hr"
        if "hour" in clean_l or "hr" in clean_l:
            return qty, "hrs"
        return qty, "units"

    return None


def _resolve_region(usage_code: str, display_region: str) -> tuple[str, str]:
    """
    Dual-strategy region resolution:
    1. Extract prefix from billing code (e.g. APS3 → ap-south-1)  — faster, more reliable
    2. Fall back to display name match  (e.g. "Asia Pacific (Mumbai)" → ap-south-1)

    This covers every billing format because:
    - Regional codes (APS3-...) resolve via prefix
    - Global codes (Global-..., TimedStorage-ByteHrs) fall back to display name
    """
    # Strategy 1: prefix from billing code
    m = re.match(r"^([A-Z][A-Z0-9]+)[-:]", usage_code, re.IGNORECASE)
    if m:
        prefix = m.group(1).upper()
        if prefix in USAGE_PREFIX_TO_REGION:
            aws = USAGE_PREFIX_TO_REGION[prefix]
            return aws, AWS_TO_GCP_REGION.get(aws, aws)

    # Strategy 2: display name
    aws = LOCATION_TO_REGION.get(display_region.lower().strip())
    if aws:
        return aws, AWS_TO_GCP_REGION.get(aws, aws)

    return "unknown", "unknown"

# =============================================================================
# SECTION 10 — TEXT-MODE PARSER  (Mode A — normal PDF)
# =============================================================================

def extract_all_services_text(pdf_path: str) -> dict[str, list[dict]]:
    """
    Single-pass state-machine parser for standard text-based AWS billing PDFs.

    State variables:
      current_svc    — which service section we're inside (or None)
      current_region — current region display string
      after_level2   — True for exactly ONE line after an L2 record is stored
      last_key       — (svc_key, index) of the last stored L2 record

    Per-line precedence:
      1. Skip blank / credit / activation lines
      2. If after_level2 → extract qty+unit (L3), clear flag
      3. Detect L0 service header (service name, no billing code)
      4. Detect L1 region header
      5. Detect section boundary (different service appears)
      6. Detect L2 billing-dimension line → extract cost, set after_level2
    """
    print(f"\n  Mode : TEXT — parsing line by line")
    records: dict[str, list[dict]] = {svc: [] for svc in SERVICES}
    current_svc    = None
    current_region = "Global"
    after_level2   = False
    last_key       = None

    with pdfplumber.open(pdf_path) as pdf:
        print(f"  Pages: {len(pdf.pages)}")

        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue

            for raw in text.split("\n"):
                line  = raw.strip()
                lower = line.lower()
                if not line:
                    continue

                # ── 1. Skip credits and noise ─────────────────────────────
                if re.search(r"\(usd\s+[\d,]+", lower):
                    after_level2 = False
                    continue
                if re.search(r"\bcredit\b|\baws activate\b|\bck discount\b", lower):
                    after_level2 = False
                    continue

                # ── 2. LEVEL 3 — qty + unit line ─────────────────────────
                # Exactly one line follows every L2 line with qty+unit.
                if after_level2 and last_key:
                    result = _extract_qty_unit(line)
                    if result:
                        qty, unit = result
                        svc_key, idx = last_key
                        records[svc_key][idx]["qty"]  = qty
                        records[svc_key][idx]["unit"] = unit
                    after_level2 = False
                    continue

                # Pre-compute these once per line (used in steps 3–6)
                level2_result = _is_level2_line(line)
                svc_in_line   = _which_service(lower)

                # ── 3. LEVEL 0 — service header ───────────────────────────
                # Service name present, no billing code → it's a section header
                if svc_in_line and not level2_result:
                    if current_svc != svc_in_line:
                        current_svc    = svc_in_line
                        current_region = "Global"
                        after_level2   = False
                    continue

                # ── 4. LEVEL 1 — region header ────────────────────────────
                if _is_region_line(line):
                    current_region = _get_region_display(line)
                    after_level2   = False
                    continue

                # ── 5. Section boundary — different service detected ───────
                # Exit current section if we see another service's name
                if current_svc and not level2_result:
                    for svc_key2, svc_def2 in SERVICES.items():
                        if svc_key2 == current_svc:
                            continue
                        if any(n in lower for n in svc_def2["svc_names"]):
                            current_svc  = None
                            after_level2 = False
                            break

                if not current_svc:
                    continue

                # ── 6. LEVEL 2 — billing-dimension line ───────────────────
                if level2_result:
                    _, usage_code, cost = level2_result

                    # OWNERSHIP CHECK: the L2 line must contain THIS service's name.
                    # Prevents Fargate/ECS/Lambda lines from polluting ECR/Redshift
                    # sections when section boundaries are missed.
                    if not any(n in lower for n in SERVICES[current_svc]["svc_names"]):
                        current_svc  = None
                        after_level2 = False
                        continue

                    aws_region, gcp_region = _resolve_region(usage_code, current_region)
                    records[current_svc].append({
                        "region":     current_region,
                        "usage_code": usage_code,
                        "cost":       cost,
                        "qty":        0.0,   # filled by next L3 line
                        "unit":       "",    # filled by next L3 line
                        "aws_region": aws_region,
                        "gcp_region": gcp_region,
                    })
                    last_key     = (current_svc, len(records[current_svc]) - 1)
                    after_level2 = True
                    continue

    # Summary
    print(f"\n  Extraction results:")
    for svc, recs in records.items():
        total = sum(r["cost"] for r in recs if r["cost"] > 0)
        mark  = "V" if recs else "-"
        print(f"  {mark}  {svc:<22} {len(recs):>3} rows   AWS Total: ${total:.4f}")

    return records

# =============================================================================
# SECTION 11 — IMAGE/REDACTED MODE PROCESSOR  (Mode B)
# =============================================================================

def process_redacted_data() -> dict[str, list[dict]]:
    """
    Processes REDACTED_DATA (Section 5):
      1. Resolves AWS + GCP regions from usage_code prefix and display name
      2. Back-calculates AWS cost = qty × AWS public list rate (Section 4)
    """
    print(f"\n  Mode : IMAGE/REDACTED — using REDACTED_DATA quantities")
    print(f"  AWS costs will be estimated from public pricing rates\n")
    records: dict[str, list[dict]] = {}

    for svc_key, entries in REDACTED_DATA.items():
        svc_recs = []
        for e in entries:
            aws_region, gcp_region = _resolve_region(e["usage_code"], e["region"])
            est_cost = back_calc_aws_cost(svc_key, e["qty"], e["unit"])
            svc_recs.append({
                "region":     e["region"],
                "usage_code": e["usage_code"],
                "cost":       est_cost,
                "qty":        e["qty"],
                "unit":       e["unit"],
                "aws_region": aws_region,
                "gcp_region": gcp_region,
                "estimated":  True,
            })
        records[svc_key] = svc_recs
        total = sum(r["cost"] for r in svc_recs)
        print(f"  V  {svc_key:<22} {len(svc_recs):>3} rows   Est. AWS: ${total:.4f}"
              f"  (qty x public rate)")

    for svc_key in SERVICES:
        if svc_key not in records:
            records[svc_key] = []

    return records

# =============================================================================
# SECTION 12 — GCP COST CALCULATORS
# =============================================================================

def _gcp_ecr(qty: float, cost: float, unit: str) -> float:
    """ECR → Artifact Registry.  $0.10/GB-Mo on both. Direct 1:1."""
    gb = qty if "gb" in unit.lower() and qty > 0 else (cost / 0.10 if cost > 0 else 0)
    return round(gb * 0.10, 4)


def _gcp_redshift(qty: float, cost: float, unit: str) -> float:
    """Redshift → BigQuery.
       GB-Mo  (storage) → $0.02/GB-Mo
       RPU-Hr (compute) → $0.04/slot-hr  (BigQuery Enterprise)
    """
    u = unit.lower()
    if "gb" in u and qty > 0:  return round(qty * 0.02, 4)
    if "rpu" in u and qty > 0: return round(qty * 0.04, 4)
    # Provisioned node-hours fallback → on-demand $5/TB (10 GB/node-hr est.)
    hrs = qty if qty > 0 else (cost / 0.64 if cost > 0 else 0)
    return round((hrs * 10 / 1024) * 5.0, 4)


def _gcp_secrets(qty: float, cost: float, unit: str) -> float:
    """Secrets Manager → Secret Manager.
       Secrets:  AWS $0.40 → GCP $0.06/secret-month
       API ops:  AWS $0.05 → GCP $0.03/10K ops
    """
    u = unit.lower()
    if "secret" in u and qty > 0:                   return round(qty * 0.06, 4)
    if ("request" in u or "api" in u) and qty > 0:  return round((qty / 10_000) * 0.03, 4)
    return round((cost / 0.40) * 0.06, 4) if cost > 0 else 0.0


def _gcp_stepfn(qty: float, cost: float, unit: str) -> float:
    """Step Functions → Cloud Workflows.
       AWS $0.025/1K transitions → GCP $0.01/1K steps (first 5K/mo free)
    """
    transitions = qty if qty > 0 else ((cost / 0.025) * 1_000 if cost > 0 else 0)
    billable    = max(0.0, transitions - 5_000)
    return round((billable / 1_000) * 0.01, 4)


def _gcp_waf(qty: float, cost: float, unit: str, usage_code: str) -> float:
    """WAF → Cloud Armor.
       WebACL:       $5/ACL-mo   → $5/policy-mo  (1:1)
       Rule:         $1/rule-mo  → $1/rule-mo    (1:1)
       Requests std: $0.60/M req → $0.75/M req
       Bot Control:  $10/mo sub  → ~$5/mo (Adaptive Protection est.)
       Anti-DDoS:    $20/mo sub  → $0 (included in Armor Standard)
    """
    code_l = usage_code.lower()
    u      = unit.lower()
    if "webacl" in code_l:
        count = qty if qty > 0 else (cost / 5.0 if cost > 0 else 0)
        return round(count * 5.0, 4)
    if "rulev2" in code_l and "request" not in code_l:
        return round((qty if qty > 0 else cost) * 1.0, 4)
    if "botcontrol" in code_l and "request" not in code_l:
        return round((qty if qty > 0 else 1) * 5.0, 4)
    if "antiddos" in code_l and "request" not in code_l:
        return 0.0   # included in Cloud Armor Standard
    if "request" in code_l or "tier" in code_l or "request" in u:
        reqs = qty if qty > 0 else ((cost / 0.60) * 1_000_000 if cost > 0 else 0)
        return round((reqs / 1_000_000) * 0.75, 4)
    return round(cost * (0.75 / 0.60), 4) if cost > 0 else 0.0


def _get_gcp_costs(svc_key: str, qty: float, cost: float,
                   unit: str, usage_code: str = "") -> float:
    if svc_key == "ECR":               return _gcp_ecr(qty, cost, unit)
    if svc_key == "Redshift":          return _gcp_redshift(qty, cost, unit)
    if svc_key == "Secrets Manager":   return _gcp_secrets(qty, cost, unit)
    if svc_key == "Step Functions":    return _gcp_stepfn(qty, cost, unit)
    if svc_key == "WAF":               return _gcp_waf(qty, cost, unit, usage_code)
    return cost

# =============================================================================
# SECTION 13 — BUILD DATAFRAMES
# =============================================================================

AWS_COLS  = ["AWS Region", "AWS Region Code", "AWS Usage Code",
             "AWS Quantity", "AWS Unit", "AWS Cost ($)", "AWS Service"]
GCP_COLS  = ["GCP Service", "GCP Region", "GCP PAYG ($)", "GCP Rate"]
COMP_COLS = ["Est. Savings ($)", "Recommendation"]
SPACERS   = [" ", "  "]
ALL_COLS  = AWS_COLS + SPACERS + GCP_COLS + COMP_COLS


def build_dataframe(svc_key: str, records: list[dict],
                    is_estimated: bool = False) -> pd.DataFrame:
    svc_def = SERVICES[svc_key]
    rows    = []

    for r in records:
        if r["cost"] < 0:   # skip credits
            continue
        qty, cost, unit = r["qty"], r["cost"], r["unit"]
        usage_code      = r["usage_code"]
        payg            = _get_gcp_costs(svc_key, qty, cost, unit, usage_code)
        savings         = round(cost - payg, 4)
        rec_label       = "GCP OK" if savings >= 0 else "AWS OK"
        if is_estimated or r.get("estimated", False):
            rec_label  += " (est.)"

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
            "Recommendation":   rec_label,
        })

    cols = AWS_COLS + GCP_COLS + COMP_COLS
    return pd.DataFrame(rows, columns=cols) if rows else pd.DataFrame(columns=cols)

# =============================================================================
# SECTION 14 — EXCEL EXPORTER
# =============================================================================

COL_WIDTHS = [28, 16, 45, 14, 18, 14, 22, 3, 3, 22, 22, 14, 28, 16, 18]
COL_FMTS   = [None, None, None, "#,##0.000000", None, "$#,##0.0000", None,
              None, None, None, None, "$#,##0.0000", None, "$#,##0.0000", None]


def _write_sheet(workbook, svc_key: str, df: pd.DataFrame,
                 is_estimated: bool = False):
    svc_def = SERVICES[svc_key]

    aws_hdr  = workbook.add_format({"bold": True, "bg_color": "#FF9900",
                                    "font_color": "white", "border": 1,
                                    "align": "center", "valign": "vcenter",
                                    "text_wrap": True})
    gcp_hdr  = workbook.add_format({"bold": True, "bg_color": "#4285F4",
                                    "font_color": "white", "border": 1,
                                    "align": "center", "valign": "vcenter",
                                    "text_wrap": True})
    cmp_hdr  = workbook.add_format({"bold": True, "bg_color": "#34A853",
                                    "font_color": "white", "border": 1,
                                    "align": "center", "valign": "vcenter",
                                    "text_wrap": True})
    spc_fmt  = workbook.add_format({"bg_color": "#FFFFFF", "border": 0})
    ttl_fmt  = workbook.add_format({"bold": True, "font_size": 13,
                                    "bg_color": "#2C3E50", "font_color": "white",
                                    "border": 1, "align": "center", "valign": "vcenter"})
    wrn_fmt  = workbook.add_format({"bold": True, "font_size": 10,
                                    "bg_color": "#E74C3C", "font_color": "white",
                                    "border": 2, "align": "center", "valign": "vcenter",
                                    "text_wrap": True})
    tot_fmt  = workbook.add_format({"bold": True, "num_format": "$#,##0.0000",
                                    "bg_color": "#FFF2CC", "border": 1,
                                    "align": "center"})
    tot_lbl  = workbook.add_format({"bold": True, "bg_color": "#FFF2CC", "border": 1})
    no_data  = workbook.add_format({"italic": True, "font_color": "#888888",
                                    "font_size": 11})
    green_bg = "#E8F5E9"; red_bg = "#FFEBEE"
    even_bg  = "#F5F5F5"; odd_bg = "#FFFFFF"

    ws = workbook.add_worksheet(svc_def["sheet_name"][:31])

    # Row 0 — Title
    ws.set_row(0, 26)
    ws.merge_range(0, 0, 0, len(ALL_COLS) - 1,
                   f"  {svc_def['sheet_name']}  |  "
                   f"{svc_def['aws_service']}  to  {svc_def['gcp_service']}  |  TCO Mapping",
                   ttl_fmt)

    hdr_row = 1

    # Row 1 — Warning banner (redacted/estimated only)
    if is_estimated:
        ws.set_row(1, 40)
        ws.merge_range(1, 0, 1, len(ALL_COLS) - 1,
                       "REDACTED / IMAGE PDF  --  AWS Cost ($) values are ESTIMATES "
                       "(Quantity x AWS public list price).  "
                       "Actual costs may differ.  "
                       "Request the unredacted original PDF for exact values.",
                       wrn_fmt)
        hdr_row = 2

    # Column headers
    ws.set_row(hdr_row, 22)
    for ci, col in enumerate(ALL_COLS):
        if col in AWS_COLS:    ws.write(hdr_row, ci, col, aws_hdr)
        elif col in GCP_COLS:  ws.write(hdr_row, ci, col, gcp_hdr)
        elif col in COMP_COLS: ws.write(hdr_row, ci, col, cmp_hdr)
        else:                  ws.write(hdr_row, ci, "", spc_fmt)

    for ci, (w, fmt_str) in enumerate(zip(COL_WIDTHS, COL_FMTS)):
        cell_fmt = workbook.add_format({"num_format": fmt_str}) if fmt_str else None
        ws.set_column(ci, ci, w, cell_fmt)

    data_start = hdr_row + 1

    if df.empty:
        ws.merge_range(data_start, 0, data_start, len(ALL_COLS) - 1,
                       f"No data found for {svc_def['aws_service']} in this PDF.", no_data)
        ws.freeze_panes(hdr_row + 1, 0)
        return

    df_out = df.copy()
    df_out[" "] = ""; df_out["  "] = ""
    df_out = df_out[ALL_COLS]

    for ri, row_data in enumerate(df_out.itertuples(index=False), start=data_start):
        savings_val = row_data[ALL_COLS.index("Est. Savings ($)")]
        row_bg      = even_bg if ri % 2 == 0 else odd_bg
        for ci, val in enumerate(row_data):
            col_name = ALL_COLS[ci]
            bg = ((green_bg if savings_val >= 0 else red_bg)
                  if col_name in ("Est. Savings ($)", "Recommendation")
                  else row_bg)
            fmt_str  = COL_FMTS[ci]
            cell_fmt = workbook.add_format({
                "num_format": fmt_str or "General",
                "bg_color": bg, "border": 1,
                "align": "center", "valign": "vcenter",
            })
            ws.write(ri, ci, val, cell_fmt)

    last_data = data_start + len(df_out) - 1
    total_row = last_data + 2
    ws.write(total_row, 0, "TOTALS", tot_lbl)

    for col_name in ["AWS Cost ($)", "GCP PAYG ($)", "Est. Savings ($)"]:
        ci  = ALL_COLS.index(col_name)
        ltr = chr(ord("A") + ci)
        r1  = data_start + 1   # Excel 1-indexed
        r2  = last_data  + 1
        ws.write_formula(total_row, ci, f"=SUM({ltr}{r1}:{ltr}{r2})", tot_fmt)

    ws.freeze_panes(hdr_row + 1, 0)
    print(f"   V  Sheet '{svc_def['sheet_name']}' -- {len(df)} rows"
          f"{'  [ESTIMATED]' if is_estimated else ''}")


def _write_methodology(workbook, is_estimated: bool = False):
    ws = workbook.add_worksheet("Methodology")
    ws.set_column("A:A", 120)

    tf = workbook.add_format({"bold": True, "font_size": 14,
                               "bg_color": "#D3D3D3", "border": 1})
    h1 = workbook.add_format({"bold": True, "font_size": 12, "font_color": "#1155cc"})
    bd = workbook.add_format({"font_size": 11, "text_wrap": True})
    mf = workbook.add_format({"font_size": 11, "font_color": "#38761d", "italic": True})
    wf = workbook.add_format({"bold": True, "font_size": 11, "font_color": "#CC0000"})

    r = 0
    ws.write(r, 0, "  FinOps TCO Mapping -- ECR | Redshift | Secrets Manager | Step Functions | WAF", tf); r += 2
    ws.write(r, 0, f"Generated : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", bd); r += 1
    mode_s = ("IMAGE/REDACTED PDF -- AWS costs are ESTIMATES (qty x public rate)"
              if is_estimated else "TEXT PDF -- AWS costs extracted directly from billing data")
    ws.write(r, 0, f"PDF Mode  : {mode_s}", bd); r += 2

    if is_estimated:
        ws.write(r, 0, "REDACTED PDF -- How AWS Costs Were Estimated", wf); r += 1
        ws.write(r, 0, "  The PDF was image-based or had its cost column blacked out.", bd); r += 1
        ws.write(r, 0, "  AWS Cost = Quantity x AWS public list price (rates below).", bd); r += 1
        ws.write(r, 0, "  Actual costs may differ due to EDP discounts, reserved pricing, or tiering.", bd); r += 1
        ws.write(r, 0, "  Fix: ask customer for the original unredacted text-based billing PDF.", bd); r += 2
        ws.write(r, 0, "AWS Public Rates Used", h1); r += 1
        ws.write(r, 0, "  ECR storage           $0.10 / GB-Mo", mf); r += 1
        ws.write(r, 0, "  Redshift storage       $0.024 / GB-Mo", mf); r += 1
        ws.write(r, 0, "  Redshift compute       $0.4375 / RPU-Hr (us-east-1)", mf); r += 1
        ws.write(r, 0, "  Secrets Manager        $0.40 / secret-month", mf); r += 1
        ws.write(r, 0, "  Secrets Manager API    $0.05 / 10K calls", mf); r += 1
        ws.write(r, 0, "  Step Functions         $0.025 / 1K transitions", mf); r += 1
        ws.write(r, 0, "  WAF Web ACL            $5.00 / month", mf); r += 1
        ws.write(r, 0, "  WAF Rule               $1.00 / month", mf); r += 1
        ws.write(r, 0, "  WAF Requests           $0.60 / million", mf); r += 2

    ws.write(r, 0, "PDF Structure -- 4-Level Hierarchy (fixed across ALL billing periods)", h1); r += 1
    ws.write(r, 0, "  L0  Service total   'EC2 Container Registry (ECR)  USD 18.00'  -- IGNORED", bd); r += 1
    ws.write(r, 0, "  L1  Region subtotal 'Asia Pacific (Mumbai)          USD 18.00'  -- IGNORED", bd); r += 1
    ws.write(r, 0, "  L2  Usage dimension 'Amazon ECR APS3-TimedStorage-ByteHrs USD 18.00' -- COST EXTRACTED", bd); r += 1
    ws.write(r, 0, "  L3  Detail line     '$0.10/GB-month ... 179.997 GB-Mo USD 18.00' -- QTY+UNIT EXTRACTED", bd); r += 2

    ws.write(r, 0, "L2 Detection Rule (structural invariant -- works for any billing period)", h1); r += 1
    ws.write(r, 0, "  Last token before 'USD <cost>' contains '-' or ':' AND length >= 7 chars.", bd); r += 1
    ws.write(r, 0, "  Short tokens (GB-Mo=5, RPU-Hr=6) are units not codes -- excluded by length.", bd); r += 1
    ws.write(r, 0, "  Ownership check: L2 line must contain the current service's own name.", bd); r += 2

    ws.write(r, 0, "L3 Qty+Unit Extraction -- Two Formats Handled", h1); r += 1
    ws.write(r, 0, "  Format A (standard): '... 179.997 GB-Mo USD 18.00' -- unit word after qty", mf); r += 1
    ws.write(r, 0, "  Format B (plain):    'state transition 3,281.00 USD 0.00' -- unit inferred from desc.", mf); r += 2

    ws.write(r, 0, "TABLE 1  ECR -> Artifact Registry", h1); r += 1
    ws.write(r, 0, "  $0.10/GB-Mo on both. GCP PAYG = GB-Mo x $0.10", mf); r += 2

    ws.write(r, 0, "TABLE 2  Redshift -> BigQuery", h1); r += 1
    ws.write(r, 0, "  GB-Mo  (storage) -> $0.02/GB-Mo", mf); r += 1
    ws.write(r, 0, "  RPU-Hr (compute) -> $0.04/slot-hr (BigQuery Enterprise)", mf); r += 2

    ws.write(r, 0, "TABLE 3  Secrets Manager -> Secret Manager", h1); r += 1
    ws.write(r, 0, "  Secrets: AWS $0.40 -> GCP $0.06 / secret-month", mf); r += 1
    ws.write(r, 0, "  API:     AWS $0.05 -> GCP $0.03 / 10K ops", mf); r += 2

    ws.write(r, 0, "TABLE 4  Step Functions -> Cloud Workflows", h1); r += 1
    ws.write(r, 0, "  AWS $0.025/1K transitions -> GCP $0.01/1K steps (first 5K/mo free)", mf); r += 2

    ws.write(r, 0, "TABLE 5  WAF -> Cloud Armor", h1); r += 1
    ws.write(r, 0, "  Web ACL:   AWS $5/ACL-mo   -> GCP $5/policy-mo  (1:1)", mf); r += 1
    ws.write(r, 0, "  Rule:      AWS $1/rule-mo  -> GCP $1/rule-mo    (1:1)", mf); r += 1
    ws.write(r, 0, "  Requests:  AWS $0.60/M req -> GCP $0.75/M req", mf); r += 1
    ws.write(r, 0, "  Bot Ctrl:  AWS $10/mo sub  -> GCP ~$5/mo (Adaptive Protection)", mf); r += 1
    ws.write(r, 0, "  AntiDDoS:  AWS $20/mo sub  -> GCP $0 (included in Cloud Armor Std)", mf); r += 2

    ws.write(r, 0, "NOTE: CUDs (1yr/3yr) not shown. These 5 services are usage-billed "
             "with no committed-use discounts on GCP.", bd)


def export_to_excel(all_dfs: dict[str, pd.DataFrame], output_path: str,
                    is_estimated: bool = False):
    if not HAS_XLSXWRITER:
        print("xlsxwriter not available -- writing CSVs instead")
        for svc_key, df in all_dfs.items():
            df.to_csv(output_path.replace(".xlsx", f"_{svc_key}.csv"), index=False)
        return

    print(f"\nBuilding Excel workbook: {output_path}")
    try:
        writer   = pd.ExcelWriter(output_path, engine="xlsxwriter")
        workbook = writer.book
        for svc_key in SERVICES:
            _write_sheet(workbook, svc_key, all_dfs[svc_key], is_estimated)
        _write_methodology(workbook, is_estimated)
        print("   V  Sheet 'Methodology'")
        workbook.close()
        print(f"\nExcel saved -> {output_path}")
    except PermissionError:
        print(f"\n'{output_path}' is open in Excel. Close it and run again.")
        sys.exit(1)

# =============================================================================
# SECTION 15 — MAIN
# =============================================================================

def main():
    print("\n" + "=" * 70)
    print("  AWS -> GCP TCO  |  ECR  Redshift  Secrets  Step Functions  WAF")
    print("  Handles: text PDFs, image PDFs, redacted PDFs, any billing period")
    print("=" * 70)

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    os.makedirs(INPUT_FOLDER,  exist_ok=True)

    # ── Find PDF ──────────────────────────────────────────────────────────────
    pdf_path = None
    for f in sorted(os.listdir(INPUT_FOLDER)):
        if f.lower().endswith(".pdf"):
            pdf_path = os.path.join(INPUT_FOLDER, f)
            break

    if not pdf_path:
        print(f"\nNo PDF in '{INPUT_FOLDER}/' -- running DEMO MODE (sample data)\n")
        output_path  = os.path.join(OUTPUT_FOLDER, "demo_TCO_Map.xlsx")
        raw          = DEMO_RECORDS
        is_estimated = False

    else:
        output_path = os.path.join(OUTPUT_FOLDER,
                                   f"{Path(pdf_path).stem}_TCO_Map.xlsx")
        print(f"\n[LIVE] PDF : {pdf_path}")

        # ── Auto-detect PDF mode ──────────────────────────────────────────────
        pdf_mode = detect_pdf_mode(pdf_path)
        print(f"  PDF type : {pdf_mode.upper()}")

        if pdf_mode == "image":
            print("""
  +--------------------------------------------------------------+
  |  IMAGE / REDACTED PDF DETECTED                               |
  |  pdfplumber extracted 0 characters from this PDF.            |
  |                                                              |
  |  Possible reasons:                                           |
  |   - PDF is a scanned image or screenshot                    |
  |   - Cost column is blacked out (redacted by customer)        |
  |                                                              |
  |  AUTOMATIC ACTION:                                           |
  |   - Reading quantities from REDACTED_DATA (Section 5)       |
  |   - Estimating AWS costs: qty x AWS public list price       |
  |   - Computing GCP PAYG from same quantities                 |
  |                                                              |
  |  For exact results: request the original unredacted PDF     |
  +--------------------------------------------------------------+
""")
            raw          = process_redacted_data()
            is_estimated = True

        else:   # text PDF
            raw          = extract_all_services_text(pdf_path)
            is_estimated = False
            total_found  = sum(len(v) for v in raw.values())
            if total_found == 0:
                print("\n  No target services found in this text PDF.")
                print("  Confirm the PDF contains ECR / Redshift / Secrets / "
                      "Step Functions / WAF billing data.")

    # ── Calculate GCP costs ───────────────────────────────────────────────────
    print(f"\n--- Calculating GCP Costs ---")
    all_dfs: dict[str, pd.DataFrame] = {}

    for svc_key in SERVICES:
        records = raw.get(svc_key, [])
        df      = build_dataframe(svc_key, records, is_estimated)
        all_dfs[svc_key] = df
        if not df.empty:
            a = df["AWS Cost ($)"].sum()
            g = df["GCP PAYG ($)"].sum()
            est_tag = " [est.]" if is_estimated else ""
            print(f"   {svc_key:<22}  {len(df):>2} rows   "
                  f"AWS ${a:>10.4f}{est_tag}  ->  GCP ${g:>10.4f}  "
                  f"({'GCP cheaper' if g < a else 'AWS cheaper' if a < g else 'Equal'})")
        else:
            print(f"   {svc_key:<22}   0 rows  (not found in this PDF)")

    # ── Export ────────────────────────────────────────────────────────────────
    print(f"\n--- Exporting to Excel ---")
    export_to_excel(all_dfs, output_path, is_estimated)

    # ── Final summary ─────────────────────────────────────────────────────────
    g_aws  = sum(df["AWS Cost ($)"].sum() for df in all_dfs.values() if not df.empty)
    g_payg = sum(df["GCP PAYG ($)"].sum() for df in all_dfs.values() if not df.empty)

    print("\n" + "=" * 70)
    if is_estimated:
        print("  NOTE: AWS Cost values marked [~] are ESTIMATES (qty x public rate)")
    print(f"  {'Service':<22} {'AWS Cost':>12}  {'GCP PAYG':>12}  {'Savings':>12}")
    print(f"  {'-' * 62}")
    for svc_key, df in all_dfs.items():
        a   = df["AWS Cost ($)"].sum() if not df.empty else 0
        g   = df["GCP PAYG ($)"].sum() if not df.empty else 0
        est = " ~" if is_estimated and not df.empty else "  "
        print(f"  {svc_key:<22} ${a:>11.4f}{est}  ${g:>11.4f}   ${a - g:>11.4f}")
    print(f"  {'-' * 62}")
    print(f"  {'GRAND TOTAL':<22} ${g_aws:>11.4f}   ${g_payg:>11.4f}   ${g_aws - g_payg:>11.4f}")
    print(f"\n  Output: {output_path}")
    print("=" * 70)
    print("\nProcess Complete!")


if __name__ == "__main__":
    main()
