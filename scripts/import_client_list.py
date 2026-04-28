from __future__ import annotations

import json
import re
import shutil
import sys
import zipfile
from datetime import date, datetime, timedelta
from pathlib import Path
from xml.etree import ElementTree as ET


SHEET_NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REL_NS = {"pr": "http://schemas.openxmlformats.org/package/2006/relationships"}
MONTHS = {
    "JAN": 1,
    "JANUARY": 1,
    "FEB": 2,
    "FEBRUARY": 2,
    "MAR": 3,
    "MARCH": 3,
    "APR": 4,
    "APRIL": 4,
    "MAY": 5,
    "JUN": 6,
    "JUNE": 6,
    "JUL": 7,
    "JULY": 7,
    "AUG": 8,
    "AUGUST": 8,
    "SEP": 9,
    "SEPTEMBER": 9,
    "OCT": 10,
    "OCTOBER": 10,
    "NOV": 11,
    "NOVEMBER": 11,
    "DEC": 12,
    "DECEMBER": 12,
}


def col_index(cell_ref: str) -> int:
    letters = "".join(ch for ch in cell_ref if ch.isalpha())
    total = 0
    for ch in letters:
        total = total * 26 + ord(ch.upper()) - ord("A") + 1
    return total - 1


def parse_scalar(value: str | None):
    if value is None:
        return ""

    value = value.strip()
    if value == "":
        return ""

    try:
        number = float(value)
    except ValueError:
        return value

    return int(number) if number.is_integer() else number


def read_xlsx(path: Path) -> dict[str, list[list[object]]]:
    with zipfile.ZipFile(path) as workbook_zip:
        names = workbook_zip.namelist()
        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in names:
            root = ET.fromstring(workbook_zip.read("xl/sharedStrings.xml"))
            for item in root.findall("a:si", SHEET_NS):
                parts = [
                    text_node.text or ""
                    for text_node in item.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t")
                ]
                shared_strings.append("".join(parts))

        workbook = ET.fromstring(workbook_zip.read("xl/workbook.xml"))
        rels = ET.fromstring(workbook_zip.read("xl/_rels/workbook.xml.rels"))
        rel_targets = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels.findall("pr:Relationship", REL_NS)}

        sheets: dict[str, list[list[object]]] = {}
        for sheet in workbook.findall("a:sheets/a:sheet", SHEET_NS):
            name = sheet.attrib["name"]
            rel_id = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
            target = rel_targets[rel_id]
            sheet_path = "xl/" + target.lstrip("/") if not target.startswith("xl/") else target
            root = ET.fromstring(workbook_zip.read(sheet_path))
            rows: list[list[object]] = []
            for row in root.findall("a:sheetData/a:row", SHEET_NS):
                values: list[object] = []
                for cell in row.findall("a:c", SHEET_NS):
                    idx = col_index(cell.attrib.get("r", "A1"))
                    while len(values) <= idx:
                        values.append("")

                    cell_type = cell.attrib.get("t")
                    value_node = cell.find("a:v", SHEET_NS)
                    if cell_type == "s" and value_node is not None:
                        value = shared_strings[int(value_node.text or "0")]
                    elif cell_type == "inlineStr":
                        value = "".join(text.text or "" for text in cell.findall(".//a:t", SHEET_NS))
                    else:
                        value = parse_scalar(value_node.text if value_node is not None else None)
                    values[idx] = value
                rows.append(values)
            sheets[name] = rows
        return sheets


def cell(row: list[object], idx: int):
    return row[idx] if idx < len(row) else ""


def text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def money(value) -> float:
    if isinstance(value, (int, float)):
        return float(value)

    raw = text(value)
    if raw == "":
        return 0.0

    raw = raw.replace(",", "").replace("PHP", "").strip()
    if not re.fullmatch(r"-?\d+(\.\d+)?", raw):
        return 0.0

    return float(raw)


def positive_money(value) -> float | None:
    amount = money(value)
    return amount if amount > 0 else None


def excel_date(value, fallback: date | None = None) -> date | None:
    if isinstance(value, (int, float)) and value > 20000:
        return (datetime(1899, 12, 30) + timedelta(days=float(value))).date()

    raw = text(value)
    if raw == "":
        return fallback

    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%b %d, %Y", "%B %d, %Y"):
        try:
            return datetime.strptime(raw, fmt).date()
        except ValueError:
            pass

    return fallback


def iso_date(value, fallback: date | None = None) -> str | None:
    parsed = excel_date(value, fallback)
    return parsed.isoformat() if parsed else None


def normalize_name(value: str) -> str:
    value = value.upper()
    value = re.sub(r"\([^)]*\)", "", value)
    value = re.sub(r"[^A-Z0-9]+", " ", value)
    return " ".join(value.split())


def sheet_month(sheet_name: str) -> date | None:
    match = re.search(r"\b([A-Z]+)\s*(20\d{2})\b", sheet_name.upper())
    if not match:
        return None

    month = MONTHS.get(match.group(1))
    if month is None:
        return None

    return date(int(match.group(2)), month, 1)


def header_indexes(row: list[object]) -> dict[str, list[int]]:
    indexes: dict[str, list[int]] = {}
    for idx, value in enumerate(row):
        label = text(value).strip().upper()
        if label:
            indexes.setdefault(label, []).append(idx)
    return indexes


def header_col(indexes: dict[str, list[int]], label: str, occurrence: int = 0) -> int | None:
    matches = indexes.get(label.upper(), [])
    if len(matches) <= occurrence:
        return None
    return matches[occurrence]


def next_id(records: list[dict]) -> int:
    return max((int(record.get("id", 0)) for record in records), default=0) + 1


def build_clients(rows: list[list[object]]) -> tuple[list[dict], dict[str, int], dict[str, int]]:
    clients: list[dict] = []
    by_name: dict[str, int] = {}
    by_pppoe: dict[str, int] = {}

    for row in rows[1:]:
        name = text(cell(row, 7))
        if not name:
            continue

        account_number = text(cell(row, 1)) or str(len(clients) + 1)
        plan = money(cell(row, 4))
        status = text(cell(row, 3)) or text(cell(row, 0)) or "Active"
        area = text(cell(row, 5))
        zone = text(cell(row, 6))
        pppoe = text(cell(row, 8))
        facebook = text(cell(row, 9)) or text(cell(row, 14))

        client = {
            "id": len(clients) + 1,
            "accountNumber": account_number,
            "dateInstalled": iso_date(cell(row, 2)),
            "status": status,
            "billingType": "Prepaid",
            "planAmount": plan,
            "area": area,
            "zone": zone,
            "name": name,
            "pppoeUsername": pppoe,
            "contact": facebook,
            "email": "",
            "facebookAccount": facebook,
            "balance": money(cell(row, 10)),
            "advance": money(cell(row, 11)),
            "bills": money(cell(row, 12)),
            "address": " ".join(part for part in (area, zone) if part),
            "latitude": None,
            "longitude": None,
            "creditLimit": 0,
            "remarks": "Imported from CLIENTS LIST sheet.",
        }
        clients.append(client)
        by_name.setdefault(normalize_name(name), client["id"])
        if pppoe:
            by_pppoe.setdefault(pppoe.upper(), client["id"])

    return clients, by_name, by_pppoe


def match_client(name: str, pppoe: str, by_name: dict[str, int], by_pppoe: dict[str, int]) -> int:
    if pppoe and pppoe.upper() in by_pppoe:
        return by_pppoe[pppoe.upper()]

    normalized = normalize_name(name)
    if normalized in by_name:
        return by_name[normalized]

    trimmed = re.sub(r"\b(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC|JANUARY|FEBRUARY|MARCH|APRIL)\b.*", "", normalized)
    trimmed = " ".join(trimmed.split())
    return by_name.get(trimmed, 0)


def import_monthly_bills(
    sheets: dict[str, list[list[object]]],
    clients: list[dict],
    by_name: dict[str, int],
    by_pppoe: dict[str, int],
) -> list[dict]:
    client_by_id = {client["id"]: client for client in clients}
    overrides: list[dict] = []
    latest_month: date | None = None
    latest_rows: list[tuple[int, dict[str, list[int]], list[object]]] = []

    for sheet_name, rows in sheets.items():
        if not re.fullmatch(r"BILLS\s+[A-Z]{3,9}20\d{2}", sheet_name.upper()):
            continue

        month = sheet_month(sheet_name)
        if month is None:
            continue

        header_row = None
        indexes = None
        for row in rows[:10]:
            possible_indexes = header_indexes(row)
            if "NAME" in possible_indexes and "BILLS" in possible_indexes:
                header_row = row
                indexes = possible_indexes
                break
        if header_row is None or indexes is None:
            continue

        name_col = header_col(indexes, "NAME")
        pppoe_col = header_col(indexes, "PPPOE")
        status_col = header_col(indexes, "STATUS")
        plan_col = header_col(indexes, "PLAN")
        bill_col = header_col(indexes, "BILLS")
        balance_col = header_col(indexes, "BALANCE")
        final_balance_col = header_col(indexes, "BALANCE", 1)
        advance_col = header_col(indexes, "ADVANCE")
        final_advance_col = header_col(indexes, "ADVANCE", 1)

        month_rows: list[tuple[int, dict[str, list[int]], list[object]]] = []
        for row in rows[rows.index(header_row) + 1 :]:
            client_name = text(cell(row, name_col if name_col is not None else -1))
            if not client_name:
                continue

            client_id = match_client(client_name, text(cell(row, pppoe_col)) if pppoe_col is not None else "", by_name, by_pppoe)
            if client_id == 0:
                continue

            plan = money(cell(row, plan_col)) if plan_col is not None else 0
            balance = money(cell(row, balance_col)) if balance_col is not None else 0
            advance = money(cell(row, advance_col)) if advance_col is not None else 0
            bill = max(0, balance + plan - advance)
            if bill <= 0 and bill_col is not None:
                bill = money(cell(row, bill_col))
            status = text(cell(row, status_col)) if status_col is not None else ""
            if bill > 0 or status.upper() in {"DC", "CUT", "DISCONNECTED"}:
                overrides.append(
                    {
                        "id": len(overrides) + 1,
                        "clientId": client_id,
                        "billingMonth": month.isoformat(),
                        "billAmount": bill,
                        "balance": balance,
                        "advance": advance,
                        "recordedAt": datetime.now().isoformat(timespec="seconds"),
                        "remarks": f"Imported from {sheet_name}.",
                    }
                )

            if plan > 0:
                client_by_id[client_id]["planAmount"] = plan

            month_rows.append((client_id, indexes, row))

        if latest_month is None or month > latest_month:
            latest_month = month
            latest_rows = month_rows

    for client_id, indexes, row in latest_rows:
        client = client_by_id[client_id]
        status_col = header_col(indexes, "STATUS")
        bill_col = header_col(indexes, "BILLS")
        balance_col = header_col(indexes, "BALANCE", 1) or header_col(indexes, "BALANCE")
        advance_col = header_col(indexes, "ADVANCE", 1) or header_col(indexes, "ADVANCE")

        status = text(cell(row, status_col)) if status_col is not None else ""
        if status:
            client["status"] = status
        if bill_col is not None:
            client["bills"] = money(cell(row, bill_col))
        if balance_col is not None:
            client["balance"] = money(cell(row, balance_col))
        if advance_col is not None:
            client["advance"] = money(cell(row, advance_col))

    return overrides


def import_cash_payments(
    sheets: dict[str, list[list[object]]],
    by_name: dict[str, int],
    by_pppoe: dict[str, int],
) -> list[dict]:
    payments: list[dict] = []
    for sheet_name, rows in sheets.items():
        if not re.fullmatch(r"CASH\s+[A-Z]{3,9}20\d{2}", sheet_name.upper()):
            continue

        default_month = sheet_month(sheet_name)
        fallback_date = default_month or date.today().replace(day=1)
        last_date = fallback_date

        for row in rows[5:]:
            paid_on = excel_date(cell(row, 3), last_date)
            if paid_on is not None:
                last_date = paid_on

            client_name = text(cell(row, 6))
            amount = positive_money(cell(row, 7))
            if not client_name or amount is None:
                continue

            client_id = match_client(client_name, "", by_name, by_pppoe)
            remarks = f"Imported from {sheet_name}"
            if client_id == 0:
                remarks += f"; unmatched client: {client_name}"

            payments.append(
                {
                    "id": len(payments) + 1,
                    "clientId": client_id,
                    "paidOn": (paid_on or fallback_date).isoformat(),
                    "amount": amount,
                    "method": "Cash",
                    "referenceNumber": "",
                    "collectedBy": "",
                    "remarks": remarks,
                }
            )
    return payments


def import_gcash_payments(
    sheets: dict[str, list[list[object]]],
    by_name: dict[str, int],
    by_pppoe: dict[str, int],
    payments: list[dict],
) -> None:
    for sheet_name, rows in sheets.items():
        if not re.fullmatch(r"GCASH\s+[A-Z]{3,9}20\d{2}", sheet_name.upper()):
            continue

        default_month = sheet_month(sheet_name)
        fallback_date = default_month or date.today().replace(day=1)
        last_date = fallback_date

        for row in rows[6:]:
            paid_on = excel_date(cell(row, 6), last_date)
            if paid_on is not None:
                last_date = paid_on

            client_name = text(cell(row, 9))
            amount = positive_money(cell(row, 11))
            if not client_name or amount is None:
                continue

            client_id = match_client(client_name, "", by_name, by_pppoe)
            remarks = f"Imported from {sheet_name}"
            if client_id == 0:
                remarks += f"; unmatched client: {client_name}"

            payments.append(
                {
                    "id": len(payments) + 1,
                    "clientId": client_id,
                    "paidOn": (paid_on or fallback_date).isoformat(),
                    "amount": amount,
                    "method": "GCash",
                    "referenceNumber": text(cell(row, 8)),
                    "collectedBy": text(cell(row, 7)),
                    "remarks": remarks,
                }
            )


def main() -> int:
    if len(sys.argv) != 3:
        print("Usage: import_client_list.py <client-list.xlsx> <billing-data.json>", file=sys.stderr)
        return 2

    workbook_path = Path(sys.argv[1])
    data_path = Path(sys.argv[2])
    if not workbook_path.exists():
        print(f"Workbook not found: {workbook_path}", file=sys.stderr)
        return 1

    existing = json.loads(data_path.read_text(encoding="utf-8")) if data_path.exists() else {}
    backup_path = data_path.with_name(f"{data_path.stem}.backup-{datetime.now():%Y%m%d-%H%M%S}{data_path.suffix}")
    if data_path.exists():
        shutil.copy2(data_path, backup_path)

    sheets = read_xlsx(workbook_path)
    if "CLIENTS LIST" not in sheets:
        print("CLIENTS LIST sheet not found.", file=sys.stderr)
        return 1

    clients, by_name, by_pppoe = build_clients(sheets["CLIENTS LIST"])
    monthly_bill_overrides = import_monthly_bills(sheets, clients, by_name, by_pppoe)
    payments = import_cash_payments(sheets, by_name, by_pppoe)
    import_gcash_payments(sheets, by_name, by_pppoe, payments)
    payments.sort(key=lambda payment: (payment["paidOn"], payment["id"]))
    for idx, payment in enumerate(payments, start=1):
        payment["id"] = idx

    data = {
        "clients": clients,
        "payments": payments,
        "expenses": [],
        "technicians": [],
        "jobs": [],
        "pppoeUsers": [],
        "trafficSamples": [],
        "userAccounts": existing.get("userAccounts") or [
            {
                "id": 1,
                "username": "admin",
                "password": "admin123",
                "role": "Admin",
                "displayName": "Administrator",
                "technicianId": None,
                "isActive": True,
            }
        ],
        "planChanges": [],
        "monthlyBillOverrides": monthly_bill_overrides,
        "settings": existing.get("settings") or {
            "companyName": "Billing System",
            "monthlyDueDay": 15,
            "smsReminderTemplate": "Hi {Name}, your balance is {Balance}.",
            "currency": "PHP",
        },
    }

    data_path.write_text(json.dumps(data, indent=2), encoding="utf-8")
    print(f"backup={backup_path}")
    print(f"clients={len(clients)}")
    print(f"payments={len(payments)}")
    print(f"monthlyBillOverrides={len(monthly_bill_overrides)}")
    print(f"unmatchedPayments={sum(1 for payment in payments if payment['clientId'] == 0)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
