from __future__ import annotations

from datetime import datetime, timedelta
from pathlib import Path
import json
import random

STATUSES = ["Success", "Failed", "Pending", "In Progress"]
REGIONS = ["Western", "Central", "Southern", "Northern", "Eastern", "North Western"]
SOURCES = ["CRM", "Mobile App", "Web Portal", "Branch", "Call Center", "Partner API"]
PRODUCTS = ["Savings", "Loan", "Credit Card", "Insurance", "Wallet", "Remittance"]
CHANNELS = ["API", "Batch", "Manual", "Webhook"]
CUSTOMER_TYPES = ["Individual", "SME", "Corporate"]
SEGMENTS = ["Mass", "Mass Affluent", "Premium", "Enterprise"]


def generate_dummy_requests(count: int = 260) -> list[dict]:
    random.seed(42)
    base_now = datetime(2026, 5, 18, 14, 30, 0)
    rows: list[dict] = []

    for index in range(1, count + 1):
        created_at = base_now - timedelta(
            days=random.randint(0, 13),
            hours=random.randint(0, 22),
            minutes=random.randint(0, 59),
        )
        status = random.choices(STATUSES, weights=[45, 18, 22, 15], k=1)[0]
        request_id = f"REQ-{created_at.strftime('%Y%m%d')}-{index:05d}"
        amount = round(random.uniform(1500, 250000), 2)
        risk_score = random.randint(1, 100)
        attempts = random.randint(0, 4) if status != "Success" else random.randint(0, 1)
        last_error = "" if status in ["Success", "Pending", "In Progress"] else random.choice([
            "CRM validation timeout",
            "Customer profile not found",
            "Duplicate mobile number",
            "External API returned 500",
            "Invalid KYC reference",
        ])
        updated_at = created_at + timedelta(minutes=random.randint(4, 180))

        rows.append({
            "Id": index,
            "RequestId": request_id,
            "CustomerId": f"CUST{random.randint(100000, 999999)}",
            "CustomerName": random.choice([
                "A. Perera", "N. Fernando", "S. Silva", "R. Jayasinghe", "M. De Silva",
                "K. Bandara", "T. Rajapaksha", "P. Wijesinghe", "D. Gunasekara", "H. Herath",
            ]),
            "CRMStatus": status,
            "CreatedAt": created_at.isoformat(sep=" "),
            "UpdatedAt": updated_at.isoformat(sep=" "),
            "Email": f"customer{index}@example.com",
            "Phone": f"+9477{random.randint(1000000, 9999999)}",
            "NIC": f"{random.randint(700000000, 999999999)}V",
            "Region": random.choice(REGIONS),
            "Source": random.choice(SOURCES),
            "Product": random.choice(PRODUCTS),
            "Channel": random.choice(CHANNELS),
            "CustomerType": random.choice(CUSTOMER_TYPES),
            "Priority": random.choice(["Low", "Medium", "High", "Critical"]),
            "Amount": amount,
            "Currency": "LKR",
            "Attempts": attempts,
            "LastError": last_error,
            "ProcessingDurationSeconds": random.randint(12, 950),
            "AgentId": f"AGT{random.randint(100, 999)}",
            "BranchCode": f"BR{random.randint(1, 180):03d}",
            "AccountNo": f"{random.randint(1000000000, 9999999999)}",
            "ConsentStatus": random.choice(["Granted", "Pending", "Rejected"]),
            "KYCStatus": random.choice(["Verified", "Pending", "Rejected", "Not Required"]),
            "RiskScore": risk_score,
            "RiskBand": "High" if risk_score >= 75 else "Medium" if risk_score >= 45 else "Low",
            "AddressLine1": f"No {random.randint(1, 250)}, Main Street",
            "City": random.choice(["Colombo", "Kandy", "Galle", "Jaffna", "Matara", "Kurunegala"]),
            "PostalCode": f"{random.randint(10000, 99999)}",
            "Country": "Sri Lanka",
            "DOB": f"{random.randint(1970, 2003)}-{random.randint(1, 12):02d}-{random.randint(1, 28):02d}",
            "Gender": random.choice(["Male", "Female", "Other"]),
            "Segment": random.choice(SEGMENTS),
            "CampaignCode": random.choice(["CMP-MAY", "CMP-DIGITAL", "CMP-RETENTION", "CMP-NEW", ""]),
            "Reference1": f"REF-{random.randint(10000, 99999)}",
            "Reference2": f"EXT-{random.randint(10000, 99999)}",
            "CallbackUrl": "https://example.com/callback",
            "Notes": "Dummy monitoring row generated for dashboard preview.",
            "MetadataJson": json.dumps({"dummy": True, "batch": random.randint(1, 12)}),
        })

    rows.sort(key=lambda row: row["CreatedAt"], reverse=True)
    return rows


def ensure_dummy_data(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        return
    path.write_text(json.dumps(generate_dummy_requests(), indent=2), encoding="utf-8")
