from __future__ import annotations

import copy
import json
import shutil
import sys
from datetime import datetime
from pathlib import Path

from import_client_list import (
    import_cash_payments,
    import_gcash_payments,
    import_monthly_bills,
    normalize_name,
    read_xlsx,
)


def main() -> int:
    if len(sys.argv) != 3:
        print("Usage: import_payment_history.py <cash-flow.xlsx> <billing-data.json>", file=sys.stderr)
        return 2

    workbook_path = Path(sys.argv[1])
    data_path = Path(sys.argv[2])
    if not workbook_path.exists():
        print(f"Workbook not found: {workbook_path}", file=sys.stderr)
        return 1

    data = json.loads(data_path.read_text(encoding="utf-8"))
    backup_path = data_path.with_name(f"{data_path.stem}.payment-history-backup-{datetime.now():%Y%m%d-%H%M%S}{data_path.suffix}")
    shutil.copy2(data_path, backup_path)

    imports_dir = data_path.parent / "imports"
    imports_dir.mkdir(exist_ok=True)
    saved_workbook = imports_dir / f"payment-history-{datetime.now():%Y%m%d-%H%M%S}-{workbook_path.name}"
    shutil.copy2(workbook_path, saved_workbook)

    sheets = read_xlsx(workbook_path)
    clients = data.get("clients", [])
    by_name = {normalize_name(client.get("name", "")): client["id"] for client in clients if client.get("name")}
    by_pppoe = {
        client.get("pppoeUsername", "").upper(): client["id"]
        for client in clients
        if client.get("pppoeUsername")
    }

    client_copy = copy.deepcopy(clients)
    monthly_bill_overrides = import_monthly_bills(sheets, client_copy, by_name, by_pppoe)
    payments = import_cash_payments(sheets, by_name, by_pppoe)
    import_gcash_payments(sheets, by_name, by_pppoe, payments)
    payments.sort(key=lambda payment: (payment["paidOn"], payment["id"]))
    for index, payment in enumerate(payments, start=1):
        payment["id"] = index

    data["payments"] = payments
    data["monthlyBillOverrides"] = monthly_bill_overrides
    data["planChanges"] = []

    data_path.write_text(json.dumps(data, indent=2), encoding="utf-8")
    print(f"backup={backup_path}")
    print(f"savedWorkbook={saved_workbook}")
    print(f"payments={len(payments)}")
    print(f"monthlyBillOverrides={len(monthly_bill_overrides)}")
    print(f"unmatchedPayments={sum(1 for payment in payments if payment['clientId'] == 0)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
