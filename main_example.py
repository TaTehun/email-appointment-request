"""Email booking bot: customers without a portal get an appointment request by email.

Portfolio example. Addresses, customer names, and file paths are placeholders.

Flow
  1. Read the shared master and keep rows that need an email request
  2. Group them by Load ID (one email per Load, every PO + quantity listed)
  3. Resolve the customer's recipients from a contact report (alias + city match)
  4. Send, then write C/F, the requested date, and the send time back to the master
  5. One result mail with the master and a Skipped file
"""
import re
from datetime import datetime, timedelta

import openpyxl

MASTER_PATH = r"\\your-server\share\Appointment Request Master.xlsx"
CONTACT_REPORT_PATH = r"\\your-server\share\Customer Contact Report.xlsx"

CUSTOMER_MAIL_CC = ["team-tm@example.com", "team-tracking@example.com"]
CUSTOMER_MAIL_BCC = ["owner@example.com"]          # hidden from the customer

TARGET_SERVICE_TYPES = {"PN", "PR", "FT"}          # rail (PN/PR) + over-the-road (FT)
PORTAL_CUSTOMERS = {"CUSTOMER_A", "CUSTOMER_B", "CUSTOMER_C"}   # booked by their own portal bots
# Master short names → the name used in the contact report
CUSTOMER_NAME_ALIASES = {"CUST_X": "CUSTOMER X DISTRIBUTION", "CUST_Y": "CUSTOMER Y TV"}

PREFERRED_TIME = "6:00 AM"
ESTIMATED_PALLET_COUNT = 21                        # fixed estimate agreed in the requirements Q&A
EMAIL_RE = re.compile(r"[^@\s;,]+@[^@\s;,]+\.[^@\s;,]+")

REQUESTED_DT, REQUESTED_TM = "1st Requested Appt.D/T", "1st Requested Appt.T/M"
SENT_DT, SENT_TM = "1st Appt.Request D/T", "1st Appt.Request T/M"


def needs_email(row, col) -> bool:
    """Service type in scope, not a portal customer, an Initial waiting at Stage 1,
    no requested date yet, and not already C/F."""
    v = lambda name: row[col[name] - 1].value
    return (v("Service Type") in TARGET_SERVICE_TYPES
            and str(v("Sold-to Name") or "").strip().upper() not in PORTAL_CUSTOMERS
            and str(v("Request Type") or "").strip().upper() == "I"
            and str(v("Process Stage") or "").strip() == "1"
            and not str(v(REQUESTED_DT) or "").strip()
            and not str(v("APPT Scheduling") or "").strip())


def resolve_recipients(customer: str, city: str, contacts_by_customer: dict) -> list[str]:
    """Contact report rows for the customer (via alias), narrowed to the ones whose address
    contains the master's city. No address on any row → all of them. Deduped, valid emails only."""
    key = str(customer or "").strip().upper()
    name = CUSTOMER_NAME_ALIASES.get(key, key)
    rows = list(contacts_by_customer.get(name, []))
    if key != name:
        rows += contacts_by_customer.get(key, [])

    master_city = (city or "").upper().rstrip(",").strip()
    # An empty city would match every address, so it never counts as a location match
    located = [r for r in rows if master_city and r["address"] and master_city in r["address"].upper()]
    if located:
        chosen = located
    elif rows and not any(r["address"] for r in rows):
        chosen = rows
    else:
        chosen = []

    to: list[str] = []
    for r in chosen:
        addr = str(r["email"] or "").strip()
        if EMAIL_RE.fullmatch(addr) and addr.lower() not in {a.lower() for a in to}:
            to.append(addr)
    return to


def requested_date(rep: dict) -> tuple[datetime | None, str | None]:
    """OTR (FT) asks for the RDD itself; rail asks for destination ETA + 1 day."""
    if str(rep["Service Type"] or "").upper() == "FT":
        rdd = rep["RDD"]
        return (rdd, None) if isinstance(rdd, datetime) else (None, f"RDD is not a single date ({rdd!r})")
    eta = rep["Rail Desti. ETA"]
    if eta is None:
        return None, "rail ETA is missing"
    return eta + timedelta(days=1), None


def po_qty_pairs(rows: list[dict]) -> list[tuple[str, str]]:
    """A grouped row holds Cust PO and SO Qty joined with ', ': split both and pair them in order."""
    pairs = []
    for row in rows:
        pos = [p.strip() for p in str(row["Cust PO"] or "").split(",") if p.strip()]
        qtys = [q.strip() for q in str(row["SO Qty"] or "").split(",")]
        pairs += [(po, qtys[i] if i < len(qtys) else "") for i, po in enumerate(pos)]
    return pairs


def build_and_send(rows: list[dict], load_id: str, sender, contacts_by_customer: dict):
    """Returns (status 'C'|'F', reason, master_updates). Updates only on success."""
    rep = rows[0]   # the Load's first row carries the shared values
    to = resolve_recipients(rep["Sold-to Name"], rep["City"], contacts_by_customer)
    if not to:
        return "F", f"no recipient email for {rep['Sold-to Name']!r} ({rep['City']!r})", None
    date, reason = requested_date(rep)
    if date is None:
        return "F", reason, None

    lines = "<br>".join(f"{po} - {qty}" for po, qty in po_qty_pairs(rows))
    body = (f"<p>Please schedule a delivery appointment for the following load.</p>"
            f"<p>Preferred Delivery Date/Time: {date:%m/%d/%Y} {PREFERRED_TIME}</p>"
            f"<p>Load ID: {load_id}</p>"
            f"<p>Customer PO / QTY:<br>{lines}</p>"
            f"<p>Estimated Pallet Count: {ESTIMATED_PALLET_COUNT}</p>"
            f"<p>Unload Type: {rep['UNLOAD TYPE'] or 'Live'}</p>")
    subject = f"Delivery Request - {rep['Sold-to Name']}, {rep['City']} {rep['Region']} - {load_id}"

    sent_at = datetime.now()   # the send time is recorded directly instead of reading Sent Items
    if not sender.send(subject=subject, body=body, to=to, cc=CUSTOMER_MAIL_CC, bcc=CUSTOMER_MAIL_BCC):
        return "F", "customer email send failed", None
    return "C", None, {REQUESTED_DT: date, REQUESTED_TM: datetime.min.time(),
                       SENT_DT: sent_at.date(), SENT_TM: sent_at.time()}


def write_back(results: dict, updates: dict) -> None:
    """Writes by header name, opened separately from the data_only read so formulas survive."""
    wb = openpyxl.load_workbook(MASTER_PATH)
    ws = wb["Master"]
    col = {str(c.value): c.column for c in ws[1] if c.value}
    formats = {REQUESTED_DT: "MM/DD/YYYY", REQUESTED_TM: "HH:MM", SENT_DT: "MM/DD/YYYY", SENT_TM: "HH:MM"}
    for r in range(2, ws.max_row + 1):
        load_id = ws.cell(row=r, column=col["Load ID"]).value
        if load_id not in results:
            continue
        ws.cell(row=r, column=col["APPT Scheduling"], value=results[load_id])
        if results[load_id] == "C" and str(ws.cell(row=r, column=col["DEPT"]).value).upper() == "PARTNER":
            ws.cell(row=r, column=col["Process Stage"], value="2")
        for name, value in updates.get(load_id, {}).items():
            ws.cell(row=r, column=col[name], value=value).number_format = formats[name]
    wb.save(MASTER_PATH)


if __name__ == "__main__":
    contacts = {
        "CUSTOMER X DISTRIBUTION": [
            {"email": "dock-dallas@example.com", "address": "100 Main St, DALLAS, TX"},
            {"email": "dock-austin@example.com", "address": "9 Elm Rd, AUSTIN, TX"},
        ],
    }
    print(resolve_recipients("CUST_X", "Dallas,", contacts))   # ['dock-dallas@example.com']
    print(po_qty_pairs([{"Cust PO": "4501, 4502", "SO Qty": "10, 20"}]))   # [('4501', '10'), ('4502', '20')]
