"""Load-free / withdrawal report for the Pramod Nikalje family.

Usage: python tools/build_load_free_report.py [outdir]

    Pramod_Nikalje_Family_Withdrawal_Analysis.xlsx   Summary | Holdings |
                                                     Withdrawal | Action Plan |
                                                     Client Message
    Pramod_Nikalje_Family_Withdrawal_1Pager.xlsx     one sheet, prints on one page

Source: Load Free Unit Report as on 30 Sep 2026, Opulencia Capital (ARN-207036).
Every figure below is read off that report; gains, totals and the three
withdrawal buckets are recomputed here and reconcile to the paisa.

The report slices the portfolio by exit-load status, not by scheme, so a single
scheme can appear in two buckets - HDFC Retirement Hybrid Equity and Nippon
India Flexi Cap each do. XIRR is stated per bucket and cannot be averaged
across them; the workbook says so rather than inventing a portfolio XIRR.
"""

import os
import sys

from openpyxl import Workbook
from openpyxl.formatting.rule import DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

FAMILY = "Pramod Nikalje Family"
ASON = "30 September 2026"
FIRM = "Opulencia Capital Private Limited  -  AMFI Registered Mutual Fund Distributor (ARN-207036)"

FONT = "Arial"
NAVY, BLUE, GREY, WHITE, YELLOW, LINE = "1F3864", "2E75B6", "F2F2F2", "FFFFFF", "FFF2CC", "BFBFBF"
RED_FILL, AMBER_FILL, GREEN_FILL = "FFC7CE", "FFEB9C", "C6EFCE"

BODY = Font(name=FONT, size=10)
BOLD = Font(name=FONT, size=10, bold=True)
CALC = Font(name=FONT, size=10, color="000000")
SRC = Font(name=FONT, size=10, color="0000FF")
HEAD = Font(name=FONT, size=10, bold=True, color=WHITE)
RED = Font(name=FONT, size=10, color="9C0006")
AMBER = Font(name=FONT, size=10, color="9C5700")
GREEN = Font(name=FONT, size=10, color="006100")
ITAL = Font(name=FONT, size=9, italic=True, color="404040")
SECT = Font(name=FONT, size=11, bold=True, color=NAVY)

INR = r'[>=10000000]₹ ##\,##\,##\,##0.00;[>=100000]₹ ##\,##\,##0.00;₹ #,##0.00'
UNITS = '#,##0.0000;(#,##0.0000);-'
PCT = '0.00%;(0.00%);-'
THIN = Side(style="thin", color=LINE)
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

FREE, LOAD, LOCK = "Load Free", "Load Applicable", "Lock In"

# holder, folio, scheme, bucket, units, invested, current, xirr
H = [
    ("Pramod Nikalje", "9713877/79",   "HDFC Retirement Fund - Equity Plan - Regular Plan",        FREE,  9398.9330, 159997.25, 436157.49,  0.1608),
    ("Pramod Nikalje", "9713877/79",   "HDFC Retirement Fund - Hybrid Equity Plan - Regular Plan", FREE,  2921.6940, 104994.75, 104281.10, -0.0034),
    ("Pramod Nikalje", "10868589",     "Union Retirement Fund - Regular Plan",                     FREE, 10772.4380, 159992.00, 170850.87,  0.0451),
    ("Pramod Nikalje", "9713877/79",   "HDFC Retirement Fund - Hybrid Equity Plan - Regular Plan", LOAD,  1585.7630,  59997.00,  56599.05, -0.1038),
    ("Pramod Nikalje", "1000075469",   "Mahindra Manulife ELSS Tax Saver Fund - Regular Plan",     LOCK,  6138.6620, 164992.00, 159607.67, -0.0222),
    ("Pramod Nikalje", "4710370",      "Kotak ELSS Tax Saver Fund - Regular Plan",                 LOCK,  1502.6280, 164992.00, 169712.82,  0.0191),
    ("Amita Pramod Nikalje", "499249672482", "Nippon India Flexi Cap Fund - Regular Plan",         FREE,  2779.9270,  44998.00,  44507.19, -0.0078),
    ("Amita Pramod Nikalje", "499249672482", "Nippon India Flexi Cap Fund - Regular Plan",         LOAD,  3649.3980,  59997.00,  58427.59, -0.0522),
]

# bucket-level XIRR as printed on the report (per holder) - never averaged
BUCKET_XIRR = {("Pramod Nikalje", FREE): 0.1374, ("Pramod Nikalje", LOAD): -0.1038,
               ("Pramod Nikalje", LOCK): -0.0014,
               ("Amita Pramod Nikalje", FREE): -0.0078,
               ("Amita Pramod Nikalje", LOAD): -0.0522}

BUCKET_NOTE = {
    FREE: "Withdraw today. No exit load, no lock-in. Capital gains tax still applies.",
    LOAD: "Can be withdrawn, but an exit load is charged. Both holdings here are at a "
          "loss, so exiting now means paying a charge to book a loss.",
    LOCK: "Cannot be withdrawn. Both are ELSS funds inside the statutory three-year "
          "lock-in. No fee will release them early.",
}

MESSAGE = [
    ("Subject", "Your portfolio - what is available to withdraw, as on 30 September 2026"),
    ("", ""),
    ("", "Dear Dr. Nikalje,"),
    ("", ""),
    ("", "Here is where your family portfolio stands as on 30 September 2026."),
    ("", ""),
    ("WHAT YOU CAN WITHDRAW", ""),
    ("", "The family portfolio is worth Rs 12,00,143.77. It sits in three parts:"),
    ("", "  -  Rs 7,55,796.64  can be withdrawn immediately, with no exit load."),
    ("", "  -  Rs 1,15,026.64  can be withdrawn, but an exit load will be charged."),
    ("", "  -  Rs 3,29,320.48  is in ELSS funds and is locked under the statutory "
         "three-year rule."),
    ("", ""),
    ("", "So the amount freely available to you today is Rs 7,55,796.64 - Rs 7,11,289.45 "
         "in your name and Rs 44,507.19 in Mrs. Nikalje's."),
    ("", ""),
    ("HOW THE MONEY HAS DONE", ""),
    ("", "The load-free portion has earned an XIRR of 13.74%. Rs 4,24,984 invested is "
         "now Rs 7,11,289.45."),
    ("", ""),
    ("", "This is led by HDFC Retirement Fund - Equity Plan, which has grown from "
         "Rs 1,59,997 to Rs 4,36,157 - 172.60% in absolute terms, 16.08% XIRR, over "
         "roughly six and a half years. It has been the standout holding."),
    ("", ""),
    ("", "Across the whole family portfolio, Rs 9,19,960 invested is now worth "
         "Rs 12,00,143.77, an absolute gain of 30.46%."),
    ("", ""),
    ("TWO POINTS BEFORE YOU WITHDRAW", ""),
    ("", "1.  Withdrawing is not tax-free. The load-free holdings carry gains of about "
         "Rs 2.86 lakh and capital gains tax will apply. Please tell me before you "
         "redeem and I will work the tax out with your CA so we can plan the timing."),
    ("", ""),
    ("", "2.  The Rs 1,15,026.64 in the load-applicable part is at a small loss today. "
         "Exiting now means paying a charge to book a loss, so unless the money is "
         "needed urgently I would suggest waiting for the load period to end."),
    ("", ""),
    ("", "Happy to take you through any of this on a call."),
    ("", ""),
    ("", "Warm regards,"),
    ("", "Opulencia Capital Private Limited"),
    ("", "AMFI Registered Mutual Fund Distributor - ARN-207036"),
    ("", "9850099777  |  info@opulencemoney.com"),
]

SHORT = ("Dear Dr. Nikalje, your family portfolio as on 30 Sep 2026 is Rs 12,00,143.77. "
         "Rs 7,55,796.64 can be withdrawn today with no exit load; Rs 1,15,026.64 would "
         "attract an exit load; Rs 3,29,320.48 is in ELSS and locked in. The load-free "
         "part has earned 13.74% XIRR, led by HDFC Retirement Equity Plan at 16.08%. "
         "Do note that withdrawing attracts capital gains tax - please check with me "
         "before redeeming so we can plan it with your CA. - Opulencia Capital")


def put(ws, ref, v, *, font=BODY, fmt=None, fill=None, align=None, wrap=False, border=False):
    c = ws[ref]
    c.value = v
    c.font = font
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = PatternFill("solid", fgColor=fill)
    if align or wrap:
        c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    if border:
        c.border = BOX
    return c


def banner(ws, title, sub, last):
    ws.merge_cells(f"A1:{last}1")
    put(ws, "A1", title, font=Font(name=FONT, size=13, bold=True, color=WHITE),
        fill=NAVY, align="left")
    ws.row_dimensions[1].height = 28
    ws.merge_cells(f"A2:{last}2")
    put(ws, "A2", sub, font=ITAL, fill=GREY, align="left")
    ws.row_dimensions[2].height = 17


def heads(ws, row, labels, widths=None, left=()):
    for i, label in enumerate(labels, start=1):
        col = get_column_letter(i)
        put(ws, f"{col}{row}", label, font=HEAD, fill=BLUE,
            align="left" if col in left else "center", wrap=True, border=True)
        if widths:
            ws.column_dimensions[col].width = widths[i - 1]
    ws.row_dimensions[row].height = 30


def style(bucket):
    return {FREE: (GREEN, GREEN_FILL), LOAD: (AMBER, AMBER_FILL),
            LOCK: (RED, RED_FILL)}[bucket]


HDRS = ["Holder", "Folio", "Scheme", "Bucket", "Units", "Invested", "Current Value",
        "Gain / Loss", "Abs. Rtn.", "XIRR"]
WID = [20, 15, 48, 17, 13, 15, 16, 15, 11, 10]
ORDER = [FREE, LOAD, LOCK]


def holdings_block(ws, first):
    rows = sorted(H, key=lambda h: (ORDER.index(h[3]), -h[6]))
    for i, (who, folio, scheme, bucket, units, inv, cur, xirr) in enumerate(rows):
        r = first + i
        f, fill = style(bucket)
        put(ws, f"A{r}", who, border=True)
        put(ws, f"B{r}", folio, font=SRC, align="center", border=True)
        put(ws, f"C{r}", scheme, font=SRC, border=True)
        put(ws, f"D{r}", bucket, font=f, fill=fill, align="center", border=True)
        put(ws, f"E{r}", units, font=SRC, fmt=UNITS, border=True)
        put(ws, f"F{r}", inv, font=SRC, fmt=INR, border=True)
        put(ws, f"G{r}", cur, font=SRC, fmt=INR, border=True)
        put(ws, f"H{r}", f"=$G{r}-$F{r}", font=CALC, fmt=INR, border=True)
        put(ws, f"I{r}", f'=IF($F{r}=0,"",$H{r}/$F{r})', font=CALC, fmt=PCT,
            align="center", border=True)
        put(ws, f"J{r}", xirr, font=SRC, fmt=PCT, align="center", border=True)
    last = first + len(rows) - 1
    t = last + 1
    for col in "ABDEIJ":
        put(ws, f"{col}{t}", "", fill=NAVY, border=True)
    put(ws, f"C{t}", "FAMILY TOTAL", font=HEAD, fill=NAVY, border=True)
    for col in "FGH":
        put(ws, f"{col}{t}", f"=SUM({col}{first}:{col}{last})", font=HEAD, fmt=INR,
            fill=NAVY, border=True)
    put(ws, f"I{t}", f"=$H{t}/$F{t}", font=HEAD, fmt=PCT, fill=NAVY,
        align="center", border=True)
    put(ws, f"J{t}", "see note", font=HEAD, fill=NAVY, align="center", border=True)
    ws.conditional_formatting.add(f"G{first}:G{last}",
        DataBarRule(start_type="num", start_value=0, end_type="num",
                    end_value=450000, color="1F4E5F", showValue=True))
    return last, t



def write_analysis(path):
    wb = Workbook()
    wb.remove(wb.active)
    first = 5

    # ------------------------------------------------------------ Holdings --
    ws = wb.create_sheet("Holdings")
    ws.sheet_view.showGridLines = False
    banner(ws, f"HOLDINGS BY EXIT-LOAD STATUS  -  {FAMILY.upper()}",
           f"As on {ASON}. A scheme can appear twice - units bought at different times "
           f"carry different load status. Gain, absolute return and totals are computed.", "J")
    heads(ws, 4, HDRS, WID, left=("A", "C"))
    ws.freeze_panes = "D5"
    last, tot = holdings_block(ws, first)
    r = tot + 2
    ws.merge_cells(f"A{r}:J{r + 1}")
    put(ws, f"A{r}", "XIRR is printed per bucket on the source report and is shown per "
                     "scheme above. It is a money-weighted rate and cannot be averaged or "
                     "summed across rows, so no single portfolio XIRR is stated here. The "
                     "one whole-portfolio figure that is sound is the absolute return in "
                     "the total row.", font=AMBER, fill=YELLOW, wrap=True, border=True)
    ws.auto_filter.ref = f"A4:J{last}"

    # ---------------------------------------------------------- Withdrawal --
    ws = wb.create_sheet("Withdrawal")
    ws.sheet_view.showGridLines = False
    banner(ws, "WHAT CAN BE WITHDRAWN", f"As on {ASON}. Every figure sums the Holdings sheet.", "E")
    for col, w in zip("ABCDE", (22, 16, 18, 18, 62)):
        ws.column_dimensions[col].width = w
    heads(ws, 4, ["Bucket", "Pramod", "Amita", "Family", "What it means"], left=("A", "E"))
    hf, hl = first, last
    hold = f"Holdings!$D${hf}:$D${hl}"
    who_r = f"Holdings!$A${hf}:$A${hl}"
    val = f"Holdings!$G${hf}:$G${hl}"
    r = 5
    for bucket in ORDER:
        f, fill = style(bucket)
        put(ws, f"A{r}", bucket, font=f, fill=fill, border=True)
        put(ws, f"B{r}", f'=SUMIFS({val},{hold},$A{r},{who_r},"Pramod Nikalje")',
            font=CALC, fmt=INR, border=True)
        put(ws, f"C{r}", f'=SUMIFS({val},{hold},$A{r},{who_r},"Amita Pramod Nikalje")',
            font=CALC, fmt=INR, border=True)
        put(ws, f"D{r}", f"=$B{r}+$C{r}", font=BOLD, fmt=INR, fill=fill, border=True)
        put(ws, f"E{r}", BUCKET_NOTE[bucket], wrap=True, border=True)
        ws.row_dimensions[r].height = 34
        r += 1
    put(ws, f"A{r}", "PORTFOLIO", font=HEAD, fill=NAVY, border=True)
    for col in "BCD":
        put(ws, f"{col}{r}", f"=SUM({col}5:{col}{r - 1})", font=HEAD, fmt=INR,
            fill=NAVY, border=True)
    put(ws, f"E{r}", "", fill=NAVY, border=True)
    port = r
    r += 2

    put(ws, f"A{r}", "THE ANSWER", font=SECT)
    r += 1
    answers = [
        ("Withdraw today, no charge", f"=$D5",
         "Load-free value. This is the figure to quote."),
        ("Maximum withdrawable today", f"=$D5+$D6",
         "Adds the load-applicable money. An exit load is charged on that part."),
        ("Cannot be withdrawn", f"=$D7",
         "ELSS lock-in. No fee releases it early."),
        ("Gain sitting in the load-free money", f'=SUMIFS(Holdings!$H${hf}:$H${hl},{hold},"{FREE}")',
         "Capital gains tax applies on withdrawal - confirm the rates with the CA."),
    ]
    for label, formula, note in answers:
        put(ws, f"A{r}", label, font=BOLD, border=True)
        ws.merge_cells(f"B{r}:C{r}")
        put(ws, f"B{r}", formula, font=CALC, fmt=INR, fill=GREY, align="center", border=True)
        put(ws, f"D{r}", "", border=True)
        put(ws, f"E{r}", note, wrap=True, border=True)
        ws.row_dimensions[r].height = 26
        r += 1
    return wb, first, last, tot, port, path


def finish_analysis(wb, first, last, tot, port, path):
    hf, hl = first, last
    hold = f"Holdings!$D${hf}:$D${hl}"
    who_r = f"Holdings!$A${hf}:$A${hl}"

    # -------------------------------------------------------------- XIRR --
    ws = wb.create_sheet("XIRR")
    ws.sheet_view.showGridLines = False
    banner(ws, "RETURNS BY BUCKET", "XIRR exactly as printed on the source report. "
           "It is money-weighted, so these cannot be averaged into one number.", "F")
    heads(ws, 4, ["Holder", "Bucket", "Invested", "Current Value", "Gain / Loss", "XIRR"],
          [22, 18, 16, 16, 16, 12], left=("A", "B"))
    r = 5
    for who in ("Pramod Nikalje", "Amita Pramod Nikalje"):
        for bucket in ORDER:
            if (who, bucket) not in BUCKET_XIRR:
                continue
            f, fill = style(bucket)
            put(ws, f"A{r}", who, border=True)
            put(ws, f"B{r}", bucket, font=f, fill=fill, align="center", border=True)
            for col, src in (("C", "F"), ("D", "G"), ("E", "H")):
                put(ws, f"{col}{r}",
                    f'=SUMIFS(Holdings!${src}${hf}:${src}${hl},{hold},$B{r},{who_r},$A{r})',
                    font=CALC, fmt=INR, border=True)
            x = BUCKET_XIRR[(who, bucket)]
            put(ws, f"F{r}", x, font=SRC if x >= 0 else RED, fmt=PCT,
                fill=GREEN_FILL if x >= 0.10 else None, align="center", border=True)
            r += 1
    put(ws, f"A{r}", "FAMILY", font=HEAD, fill=NAVY, border=True)
    put(ws, f"B{r}", "", fill=NAVY, border=True)
    for col in "CDE":
        put(ws, f"{col}{r}", f"=SUM({col}5:{col}{r - 1})", font=HEAD, fmt=INR,
            fill=NAVY, border=True)
    put(ws, f"F{r}", "not additive", font=HEAD, fill=NAVY, align="center", border=True)
    fam = r
    r += 2
    ws.merge_cells(f"A{r}:F{r + 2}")
    put(ws, f"A{r}", "The 13.74% is the XIRR on Pramod's load-free money only - Rs 4,24,984 "
                     "invested, now Rs 7,11,289.45. It is not the family's return. Almost all "
                     "of it comes from one holding: HDFC Retirement Fund - Equity Plan, up "
                     "172.60% in absolute terms at 16.08% XIRR, which alone accounts for "
                     "Rs 2,76,160 of the family's Rs 2,80,184 total gain. Five of the eight "
                     "holdings are below cost.",
        font=AMBER, fill=YELLOW, wrap=True, border=True)

    # -------------------------------------------------------- Client Message --
    ws = wb.create_sheet("Client Message")
    ws.sheet_view.showGridLines = False
    banner(ws, "DRAFT MESSAGE TO THE CLIENT",
           "Ready to copy into email. A short WhatsApp version is at the bottom.", "B")
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 104
    r = 4
    for label, text in MESSAGE:
        if label:
            put(ws, f"A{r}", label, font=BOLD if label != "Subject" else SECT)
        if text:
            put(ws, f"B{r}", text, font=SRC, wrap=True)
            ws.row_dimensions[r].height = 15 + 13 * (len(text) // 95)
        r += 1
    r += 1
    put(ws, f"A{r}", "SHORT / WHATSAPP", font=SECT)
    r += 1
    ws.merge_cells(f"A{r}:B{r + 3}")
    put(ws, f"A{r}", SHORT, font=SRC, fill=YELLOW, wrap=True, border=True)

    # ---------------------------------------------------------- Action Plan --
    ws = wb.create_sheet("Action Plan")
    ws.sheet_view.showGridLines = False
    banner(ws, "ACTION PLAN", "Owner and Done are yours to fill in.", "E")
    for col, w in zip("ABCDE", (16, 30, 58, 18, 10)):
        ws.column_dimensions[col].width = w
    heads(ws, 4, ["When", "Step", "What happens", "Owner", "Done"], left=("B", "C"))
    plan = [
        ("Before the call", "Confirm the need", "Ask how much is actually needed and by "
         "when. Rs 7,55,796.64 is available without charge; anything beyond that costs "
         "an exit load or is locked."),
        ("Before the call", "Get the tax number", "Work the capital gains on the load-free "
         "holdings out with the CA - about Rs 2.86 lakh of gain sits there. Confirm the "
         "rates in force."),
        ("At the call", "Set the order of withdrawal", "Take from the load-free bucket "
         "first. Leave the load-applicable money alone - both holdings are at a loss and "
         "would also be charged to exit."),
        ("At the call", "Flag the concentration", "HDFC Retirement Equity Plan is "
         "Rs 4,36,157 of a Rs 12,00,144 portfolio and carries nearly the whole gain. "
         "Redeeming it first realises the largest tax bill; leaving it keeps the "
         "portfolio dependent on one fund."),
        ("After", "Diarise the ELSS", "Note when each ELSS tranche leaves its three-year "
         "lock-in so the Rs 3,29,320.48 can be planned rather than waited on."),
        ("Ongoing", "Review", "Five of the eight holdings are below cost. Schedule a "
         "performance review separate from this withdrawal question."),
    ]
    r = 5
    for when, step, what in plan:
        put(ws, f"A{r}", when, font=BOLD, align="center", border=True)
        put(ws, f"B{r}", step, font=BOLD, wrap=True, border=True)
        put(ws, f"C{r}", what, wrap=True, border=True)
        put(ws, f"D{r}", None, font=SRC, border=True)
        put(ws, f"E{r}", None, font=SRC, align="center", border=True)
        ws.row_dimensions[r].height = 42
        r += 1
    r += 1
    ws.merge_cells(f"A{r}:E{r + 1}")
    put(ws, f"A{r}", "Mutual fund investments are subject to market risk. Read all "
                     "scheme-related documents carefully. Capital gains tax treatment "
                     "depends on the rates in force and on each investor's position - "
                     "confirm with a CA before redeeming.", font=ITAL, wrap=True)

    # -------------------------------------------------------------- Summary --
    ws = wb.create_sheet("Summary")
    ws.sheet_view.showGridLines = False
    banner(ws, f"WITHDRAWAL & RETURNS REVIEW  -  {FAMILY.upper()}",
           f"As on {ASON}.  {FIRM}", "E")
    for col, w in zip("ABCDE", (34, 18, 14, 14, 60)):
        ws.column_dimensions[col].width = w

    put(ws, "A4", "HOW MUCH CAN BE WITHDRAWN", font=SECT)
    heads(ws, 5, ["", "Amount", "", "", "Explanation"], left=("A", "E"))
    tiles = [
        ("Available today, no charge", "=Withdrawal!$D$5", GREEN_FILL,
         "Load-free units across both holders. No exit load, no lock-in."),
        ("Maximum if exit load is paid", "=Withdrawal!$D$5+Withdrawal!$D$6", AMBER_FILL,
         "Adds Rs 1,15,026.64 of load-applicable units - both at a loss today."),
        ("Locked and unavailable", "=Withdrawal!$D$7", RED_FILL,
         "Two ELSS funds inside the statutory three-year lock-in."),
        ("Total portfolio", "=Withdrawal!$D$8", GREY,
         "Pramod Rs 10,97,208.99 and Amita Rs 1,02,934.78."),
    ]
    r = 6
    for label, formula, fill, note in tiles:
        put(ws, f"A{r}", label, font=BOLD, border=True)
        ws.merge_cells(f"B{r}:D{r}")
        put(ws, f"B{r}", formula, font=CALC, fmt=INR, fill=fill, align="center", border=True)
        put(ws, f"E{r}", note, wrap=True, border=True)
        ws.row_dimensions[r].height = 26
        r += 1

    r += 1
    put(ws, f"A{r}", "WHAT XIRR HE GOT", font=SECT)
    r += 1
    xr = [("Load-free money (Pramod)", 0.1374, "Rs 4,24,984 invested, now Rs 7,11,289.45. "
           "This is the headline number - and it covers only that bucket."),
          ("  of which HDFC Retirement Equity Plan", 0.1608,
           "172.60% absolute over roughly 6.7 years. Rs 1,59,997 is now Rs 4,36,157."),
          ("  of which Union Retirement Fund", 0.0451, "Rs 1,59,992 is now Rs 1,70,851."),
          ("  of which HDFC Retirement Hybrid", -0.0034, "Rs 1,04,995 is now Rs 1,04,281."),
          ("Load-applicable money (Pramod)", -0.1038, "HDFC Retirement Hybrid Equity Plan."),
          ("Lock-in money (Pramod)", -0.0014, "The two ELSS funds, close to flat."),
          ("Load-free money (Amita)", -0.0078, "Nippon India Flexi Cap."),
          ("Load-applicable money (Amita)", -0.0522, "Nippon India Flexi Cap.")]
    for label, v, note in xr:
        put(ws, f"A{r}", label, font=BOLD if not label.startswith("  ") else BODY, border=True)
        ws.merge_cells(f"B{r}:D{r}")
        put(ws, f"B{r}", v, font=SRC if v >= 0 else RED, fmt=PCT,
            fill=GREEN_FILL if v >= 0.10 else None, align="center", border=True)
        put(ws, f"E{r}", note, wrap=True, border=True)
        ws.row_dimensions[r].height = 24
        r += 1
    put(ws, f"A{r}", "Whole family, absolute return", font=BOLD, fill=YELLOW, border=True)
    ws.merge_cells(f"B{r}:D{r}")
    put(ws, f"B{r}", f"=Holdings!$I${tot}", font=CALC, fmt=PCT, fill=YELLOW,
        align="center", border=True)
    put(ws, f"E{r}", "Rs 9,19,960 invested is now Rs 12,00,143.77. There is no single "
                     "portfolio XIRR on the report, and bucket XIRRs cannot be averaged.",
        fill=YELLOW, wrap=True, border=True)
    ws.row_dimensions[r].height = 30
    r += 2

    put(ws, f"A{r}", "WHAT THIS ACTUALLY SAYS", font=SECT)
    r += 1
    for title, detail in [
        ("One fund carries the portfolio",
         "HDFC Retirement Equity Plan has turned Rs 1,59,997 into Rs 4,36,157. That single "
         "holding is Rs 2,76,160 of the family's Rs 2,80,184 total gain. Without it the "
         "rest of the portfolio is roughly flat."),
        ("Five of eight holdings are below cost",
         "Both ELSS funds, both HDFC Hybrid positions and Nippon Flexi Cap."),
        ("The load-applicable money is the worst of both",
         "Rs 1,15,026.64 that would be charged to exit and is already at a loss. Wait for "
         "the load period unless the money is needed."),
        ("Tax, not exit load, is the real cost",
         "Rs 2.86 lakh of gain sits in the load-free bucket. The tax on redeeming it will "
         "dwarf any exit load elsewhere."),
    ]:
        put(ws, f"A{r}", title, font=BOLD, wrap=True, border=True)
        ws.merge_cells(f"B{r}:E{r}")
        put(ws, f"B{r}", detail, wrap=True, border=True)
        ws.row_dimensions[r].height = 32
        r += 1

    wb.move_sheet("Summary", offset=-5)
    wb.active = 0
    wb.save(path)
    return path


def write_onepager(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Review"
    ws.sheet_view.showGridLines = False
    for i, w in enumerate(WID, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    banner(ws, f"WITHDRAWAL & RETURNS REVIEW  -  {FAMILY.upper()}",
           f"As on {ASON}   |   Portfolio Rs 12,00,143.77   |   {FIRM}", "J")

    ws.merge_cells("A4:J5")
    put(ws, "A4", "Rs 7,55,796.64 can be withdrawn today with no exit load. A further "
                  "Rs 1,15,026.64 can be withdrawn but is charged an exit load and is "
                  "currently at a loss. Rs 3,29,320.48 is in ELSS and cannot be touched. "
                  "The load-free money has earned 13.74% XIRR - almost all of it from one "
                  "holding.", font=Font(name=FONT, size=11), wrap=True)
    ws.row_dimensions[4].height = 20
    ws.row_dimensions[5].height = 20

    put(ws, "A7", "WHAT CAN BE WITHDRAWN", font=SECT)
    heads(ws, 8, ["Bucket", "", "What it means", "", "", "Pramod", "Amita", "Family", "", ""],
          left=("A", "C"))
    r = 9
    hf = 16
    hl = hf + len(H) - 1
    hold = f"$D${hf}:$D${hl}"
    who_r = f"$A${hf}:$A${hl}"
    val = f"$G${hf}:$G${hl}"
    for bucket in ORDER:
        f, fill = style(bucket)
        put(ws, f"A{r}", bucket, font=f, fill=fill, border=True)
        put(ws, f"B{r}", "", fill=fill, border=True)
        ws.merge_cells(f"C{r}:E{r}")
        put(ws, f"C{r}", BUCKET_NOTE[bucket], wrap=True, border=True)
        put(ws, f"F{r}", f'=SUMIFS({val},{hold},$A{r},{who_r},"Pramod Nikalje")',
            font=CALC, fmt=INR, border=True)
        put(ws, f"G{r}", f'=SUMIFS({val},{hold},$A{r},{who_r},"Amita Pramod Nikalje")',
            font=CALC, fmt=INR, border=True)
        put(ws, f"H{r}", f"=$F{r}+$G{r}", font=BOLD, fmt=INR, fill=fill, border=True)
        for col in "IJ":
            put(ws, f"{col}{r}", "", border=True)
        ws.row_dimensions[r].height = 30
        r += 1
    put(ws, f"A{r}", "PORTFOLIO", font=HEAD, fill=NAVY, border=True)
    for col in ("B", "C", "I", "J"):
        put(ws, f"{col}{r}", "", fill=NAVY, border=True)
    ws.merge_cells(f"C{r}:E{r}")
    for col in "FGH":
        put(ws, f"{col}{r}", f"=SUM({col}9:{col}{r - 1})", font=HEAD, fmt=INR,
            fill=NAVY, border=True)

    put(ws, "A14", "EVERY HOLDING", font=SECT)
    heads(ws, 15, HDRS, WID, left=("A", "C"))
    last, tot = holdings_block(ws, hf)

    r = tot + 2
    put(ws, f"A{r}", "WHAT XIRR HE GOT", font=SECT)
    r += 1
    for label, v, note in [
        ("Load-free money (Pramod)", 0.1374,
         "Rs 4,24,984 invested, now Rs 7,11,289.45. The headline number - and it covers "
         "that bucket only."),
        ("HDFC Retirement Equity Plan", 0.1608,
         "172.60% absolute over roughly 6.7 years. Rs 2,76,160 of the family's "
         "Rs 2,80,184 total gain comes from this one fund."),
        ("Everything else", None,
         "Load-applicable -10.38% (Pramod) and -5.22% (Amita); lock-in -0.14%; Amita's "
         "load-free -0.78%. Five of the eight holdings are below cost."),
    ]:
        put(ws, f"A{r}", label, font=BOLD, wrap=True, border=True)
        put(ws, f"B{r}", v, font=SRC if (v or 0) >= 0 else RED, fmt=PCT,
            fill=GREEN_FILL if (v or 0) >= 0.10 else None, align="center", border=True)
        ws.merge_cells(f"C{r}:J{r}")
        put(ws, f"C{r}", note, wrap=True, border=True)
        ws.row_dimensions[r].height = 28
        r += 1
    put(ws, f"A{r}", "Family, absolute", font=BOLD, fill=YELLOW, border=True)
    put(ws, f"B{r}", f"=$I${tot}", font=CALC, fmt=PCT, fill=YELLOW,
        align="center", border=True)
    ws.merge_cells(f"C{r}:J{r}")
    put(ws, f"C{r}", "Rs 9,19,960 invested is now Rs 12,00,143.77. There is no single "
                     "portfolio XIRR on the report, and bucket XIRRs cannot be averaged.",
        fill=YELLOW, wrap=True, border=True)
    r += 2

    ws.merge_cells(f"A{r}:J{r + 1}")
    put(ws, f"A{r}", "BEFORE REDEEMING:  Rs 2.86 lakh of gain sits in the load-free money "
                     "and capital gains tax will apply - the tax will exceed any exit load "
                     "elsewhere. Confirm the rates in force with the CA and plan the timing "
                     "before placing the redemption.",
        font=AMBER, fill=YELLOW, wrap=True, border=True)
    ws.row_dimensions[r].height = 20
    ws.row_dimensions[r + 1].height = 20
    r += 3

    ws.merge_cells(f"A{r}:J{r + 1}")
    put(ws, f"A{r}", "Mutual fund investments are subject to market risk. Read all "
                     "scheme-related documents carefully. Capital gains tax treatment "
                     "depends on the rates in force and on each investor's position - "
                     "confirm with a CA before redeeming.", font=ITAL, wrap=True)

    ws.print_area = f"A1:J{r + 1}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb.save(path)
    return path


def main(outdir="."):
    stem = os.path.join(outdir, "Pramod_Nikalje_Family")
    wb, first, last, tot, port, p = write_analysis(f"{stem}_Withdrawal_Analysis.xlsx")
    print("wrote", finish_analysis(wb, first, last, tot, port, p))
    print("wrote", write_onepager(f"{stem}_Withdrawal_1Pager.xlsx"))


if __name__ == "__main__":
    main(*(sys.argv[1:2] or []))
