# Email Appointment Request

A Python bot that requests delivery appointments by email for customers that have no booking portal.
It reads pending loads from a shared Excel master, writes one request per load to the right customer contacts,
and records the result back in the master.

> **Visual overview:** [project-overview.html](project-overview.html)

---

## The team system

This repo is one part of a delivery-appointment automation system built by a small team.
Every program connects only through one shared Excel master file: no program calls another directly.

```
① Request intake      4 request mailboxes → shared master            (mine)
② Date distribution   Available RDD, max 4 bookings/day per customer+DC (mine)
③ Booking             one bot per customer portal, or by email        (mine: Retailer A, Retailer B, email bot · team: 5 portals)
⑤ Confirmation        confirmation mails → master                     (team)
⑦ Upload              confirmed appointments → TMS / WMS              (mine: TMS uploader · team: WMS)
```

| Part | Owner | Repo |
|------|-------|------|
| Request Tracker (parent) + Intake + Distribution | Taehun Jang | [appointment-request-pipeline](https://github.com/TaTehun/appointment-request-pipeline) |
| Portal booking bot, Retailer A | Taehun Jang | [delivery-appointment-scheduler](https://github.com/TaTehun/delivery-appointment-scheduler) |
| Email booking bot | Taehun Jang | [email-appointment-request](https://github.com/TaTehun/email-appointment-request) |
| Portal booking bot, Retailer B (pre-production) | Taehun Jang | [dock-appointment-booking](https://github.com/TaTehun/dock-appointment-booking) |
| TMS uploader (pre-production) | Taehun Jang | [tms-appointment-uploader](https://github.com/TaTehun/tms-appointment-uploader) |
| 5 other portal bots, overseas/truck email bots, confirmation polling, WMS upload | Teammates | private |

---

## What it does

1. **Filters** the master: rail or OTR service type, not a portal customer, an Initial at Process Stage `1`, no requested date yet,
   APPT Scheduling blank.
2. **Groups** rows by Load ID; one email per Load lists every PO with its quantity.
3. **Resolves recipients** from the customer contact report: short name → alias → contacts whose address contains the load's city.
4. **Computes the requested date**: rail = destination ETA + 1 day, OTR = RDD. Preferred time 6:00 AM, estimated pallets 21, unload type from the master.
5. **Sends and writes back** APPT Scheduling `C`/`F`, the requested date/time, and the actual send date/time; Partner rows move to Process Stage `2`.
6. **Mails one result** with the master and a Skipped file listing each failed Load and why.

A BOL step (ERP screen automation through `win32com` that attaches the bill of lading) was built and is paused until the business re-enables it.

---

## Project Structure (private repo)

```
email-appointment-request/
├── main.py                 # filters, grouping, mail build, write-back
├── email_pipeline/         # SMTP sender, POP3 (BOL mail, paused)
├── utils/                  # paths, logging, dates, ERP BOL automation (paused)
└── run.bat
```

This public repo holds `main_example.py` and `project-overview.html`.

```
python main_example.py   # recipient-matching and PO/qty pairing demo
```

---

## Changelog

- **v1.2**: Recipients from the contact report (alias + city), team CC/BCC; PO/qty list in the body; OTR (FT) loads included; result mail with Skipped file.
- **v1.1**: BOL pipeline (ERP lookup → mail attachment), now paused.
- **v1.0**: Master filtering, mail assembly, SMTP send, master write-back.
