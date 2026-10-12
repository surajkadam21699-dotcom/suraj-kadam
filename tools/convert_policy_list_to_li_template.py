#!/usr/bin/env python3
"""Convert an insurer policy-list export into the LI bulk-upload template format.

The template (Li_template_*.xls) carries an "Enter data" sheet with 22 fixed
columns, an "Instructions" sheet and a "Sheet3" lookup sheet holding the
drop-down lists. This script writes the mapped rows into a copy of that very
template so the sheets, lists and column order stay exactly as the upload
portal expects.

Usage:
  python3 convert_policy_list_to_li_template.py \
      --source policyList_export.xlsx --template Li_template.xls --out filled.xls
"""
import argparse
import datetime
import re

import openpyxl
import xlrd
import xlwt
from xlutils.copy import copy as xl_copy

SALUTATIONS = {"mr", "mrs", "ms", "miss", "dr", "shri", "smt", "m/s", "mst"}

# Scheme name -> template "Scheme Sub-Type" drop-down value.
SUBTYPE_RULES = [
    (r"click\s*2?\s*protect", "Term Plan"),
    (r"pension|guaranteed pension", "Pension Plan"),
    (r"sampoorn nivesh|progrowth|unit linked|ulip", "Unit Linked Plan"),
    (r"super income|money back", "Money Back Plan"),
    (r"sanchay|endowment|savings", "Endowment"),
]

MODE_BY_MONTHS = {0: "Yearly", 6: "Half Yearly", 3: "Quarterly", 1: "Monthly (ECS)"}


def strip_salutation(name):
    parts = str(name or "").strip().split()
    while parts and parts[0].lower().strip(".") in SALUTATIONS:
        parts.pop(0)
    return " ".join(p.capitalize() if p.isupper() else p for p in parts)


def parse_date(value):
    if value in (None, ""):
        return None
    if isinstance(value, datetime.datetime):
        return value.date()
    if isinstance(value, datetime.date):
        return value
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
        try:
            return datetime.datetime.strptime(str(value).strip(), fmt).date()
        except ValueError:
            continue
    return None


def sub_type(scheme_name):
    low = str(scheme_name or "").lower()
    for pattern, value in SUBTYPE_RULES:
        if re.search(pattern, low):
            return value
    return ""


def paying_mode(start, next_due):
    """Infer the mode from the issuance-to-next-due gap; blank when unclear.

    Annual policies in the export often have a due date a few days off the
    issuance anniversary, so each candidate period is matched with a
    tolerance instead of requiring an exact multiple.
    """
    if not start or not next_due:
        return "", "premium paying mode to be confirmed"
    days = (next_due - start).days
    if days <= 0:
        return "", "premium paying mode to be confirmed"
    for months, label, tol in ((12, "Yearly", 20), (6, "Half Yearly", 10),
                               (3, "Quarterly", 6), (1, "Monthly (ECS)", 3)):
        period = months * 30.4375
        k = round(days / period)
        if k >= 1 and abs(days - k * period) <= tol:
            return label, ""
    return "", "premium paying mode to be confirmed"


def read_source(path):
    ws = openpyxl.load_workbook(path, data_only=True)["data1"]
    header = [str(c.value or "").strip() for c in next(ws.iter_rows(max_row=1))]
    idx = {name: i for i, name in enumerate(header)}
    out = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[idx["Policy No"]]:
            continue
        out.append({k: row[i] for k, i in idx.items()})
    return out


def map_row(rec, company, sub_broker_code):
    start = parse_date(rec.get("Issuance Date"))
    nxt = parse_date(rec.get("Next Premium Due Date"))
    mode, mode_note = paying_mode(start, nxt)
    notes = [n for n in (str(rec.get("Status") or "").strip(), mode_note) if n]
    return [
        strip_salutation(rec.get("Customer Name")),      # Policy Holder Name*
        "",                                              # PAN*
        sub_broker_code,                                 # Sub-Broker Code*
        str(rec.get("Policy No") or "").strip(),         # Policy Number*
        start,                                           # Start Date*
        strip_salutation(rec.get("Customer Name")),      # Proposer Name
        rec.get("Sum Assured"),                          # Risk Coverage*
        rec.get("Premium Installment"),                  # Premium Amount*
        mode,                                            # Premium Paying Mode*
        "",                                              # Premium Paying Term*
        "",                                              # Policy Term*
        None,                                            # DOB*
        "; ".join(notes),                                # Remarks
        "",                                              # Nominee Name
        "",                                              # Relation
        sub_type(rec.get("Policy Name")),                # Scheme Sub-Type*
        company,                                         # Insurance Company*
        str(rec.get("Policy Name") or "").strip(),       # Scheme Name*
        "",                                              # Cheque Number
        "",                                              # Proposal Number
        start,                                           # DOC*
        "",                                              # Recommended By
    ]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", required=True)
    ap.add_argument("--template", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--company", default="HDFC Life Insurance Co. Ltd",
                    help="must match the Insurance Company drop-down in Sheet3")
    ap.add_argument("--sub-broker-code", default="")
    args = ap.parse_args()

    src = read_source(args.source)

    rb = xlrd.open_workbook(args.template, formatting_info=True,
                            ignore_workbook_corruption=True)
    companies = {rb.sheet_by_name("Sheet3").cell_value(r, 3)
                 for r in range(2, rb.sheet_by_name("Sheet3").nrows)}
    if args.company not in companies:
        raise SystemExit(f"company {args.company!r} is not in the template drop-down")

    wb = xl_copy(rb)
    ws = wb.get_sheet(0)
    date_style = xlwt.easyxf(num_format_str="DD-MM-YYYY")
    plain = xlwt.easyxf()

    blanks = 0
    for r, rec in enumerate(src, start=1):
        for c, value in enumerate(map_row(rec, args.company, args.sub_broker_code)):
            if isinstance(value, datetime.date):
                ws.write(r, c, value, date_style)
            elif value in (None, ""):
                ws.write(r, c, "", plain)
                blanks += 1
            else:
                ws.write(r, c, value, plain)
    # clear the template's note block, now overlapped by data rows
    for r in range(len(src) + 1, max(len(src) + 1, 11)):
        for c in range(22):
            ws.write(r, c, "", plain)

    wb.save(args.out)
    print(f"{len(src)} policies written to {args.out} ({blanks} blank cells)")


if __name__ == "__main__":
    main()
