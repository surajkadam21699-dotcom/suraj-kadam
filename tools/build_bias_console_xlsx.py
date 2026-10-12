"""Build the Nifty Bias Console as a working Excel workbook.

Usage: python tools/build_bias_console_xlsx.py [output.xlsx]

The spreadsheet twin of the published console. Four weighted layers of
evidence - macro, global cues and GIFT Nifty, derivatives positioning, price
action - each a stack of indicators you set from a dropdown. The workbook
scores them, weights the layers, and runs the same confluence test: do the
layers that have a view actually agree?

Sheets: Console | Guide | Key

Every option list is scored +2 / +1 / 0 / -1 / -2 in that order, so a reading's
score is 3 - MATCH(choice, its five options, 0). Layer score is the mean of its
indicators scaled to -100..+100; the composite re-weights over the layers that
have been answered, so a half-filled sheet is not silently diluted.

No live market data ships with this file and none can: the selectors open on a
worked example and have to be replaced with today's readings from NSE, a
terminal, or NSE IX for GIFT Nifty.
"""

import sys

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from chart_fix import fix_chart

FONT = "Arial"
NAVY, BLUE, GREY, WHITE, SUNK, LINE = "17323E", "1B4D63", "F2F4F6", "FFFFFF", "E8ECEF", "C8D2D8"
BULL, BULL_BG = "15704A", "DCEFE4"
BEAR, BEAR_BG = "A32B22", "F7E1DE"
FLAT, FLAT_BG = "8A6410", "F7EDD6"

BODY = Font(name=FONT, size=10)
BOLD = Font(name=FONT, size=10, bold=True)
CALC = Font(name=FONT, size=10, color="000000")
PICK = Font(name=FONT, size=10, color="0000FF")
HEAD = Font(name=FONT, size=10, bold=True, color=WHITE)
SMALL = Font(name=FONT, size=9, color="4C5A66")
ITAL = Font(name=FONT, size=9, italic=True, color="4C5A66")
SECT = Font(name=FONT, size=11, bold=True, color=NAVY)
HUGE = Font(name=FONT, size=26, bold=True, color=NAVY)
BIG = Font(name=FONT, size=14, bold=True)

NUM0 = '+0;-0;0'
PCT0 = '0%'
THIN = Side(style="thin", color=LINE)
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# Every option list runs +2, +1, 0, -1, -2 in order - the score formula depends on it.
LAYERS = [
    ("macro", "Macro & fundamental", 0.20,
     "The regime. Moves over quarters - re-read on data releases, not daily.", [
        ("Real GDP growth (YoY)",
         "MoSPI quarterly print. Above 7% is among the fastest of any large economy and supports earnings expansion.",
         ["Above 7.5% - strong", "6.5 to 7.5% - healthy", "5.5 to 6.5% - moderating",
          "4.5 to 5.5% - slowing", "Below 4.5% - weak"]),
        ("CPI inflation vs RBI band",
         "RBI targets 4% with a +/-2% tolerance band. Above 6% forces a hawkish stance and compresses valuations.",
         ["Below 4% - supportive", "4 to 5% - comfortable", "5 to 6% - upper band",
          "6 to 7% - above tolerance", "Above 7% - policy pressure"]),
        ("RBI stance",
         "Direction beats level. A cutting cycle lifts rate-sensitive sectors well before the cuts reach earnings.",
         ["Cutting / accommodative", "On hold, dovish tone", "On hold, neutral",
          "On hold, hawkish tone", "Hiking / tightening"]),
        ("Composite PMI",
         "Above 50 is expansion. India has run among the highest composite readings globally; treat 55+ as genuine strength.",
         ["Above 58 - very strong", "55 to 58 - strong", "52 to 55 - expanding",
          "50 to 52 - barely expanding", "Below 50 - contracting"]),
        ("Brent crude",
         "India imports roughly 85% of its crude. Falling crude eases the current account, inflation and input costs at once.",
         ["Below $65 - strong tailwind", "$65 to $75 - supportive", "$75 to $85 - neutral",
          "$85 to $95 - headwind", "Above $95 - serious drag"]),
        ("USD/INR trend",
         "A depreciating rupee erodes FII returns in dollar terms and can turn a flat index into a loss for a foreign holder.",
         ["Appreciating", "Stable", "Mild depreciation", "Steady depreciation", "Sharp depreciation"]),
        ("Nifty P/E vs long-run mean",
         "The long-run mean sits near 20-22x. Rich valuations do not time a top, but they cap how far good news can carry.",
         ["Below 18x - cheap", "18 to 21x - fair", "21 to 24x - full",
          "24 to 27x - rich", "Above 27x - stretched"]),
        ("Nifty earnings momentum",
         "Consensus EPS revisions. Upgrades turn a liquidity rally into a durable one; without them a re-rating is borrowed.",
         ["Broad upgrades", "Mild upgrades", "Flat", "Mild downgrades", "Broad downgrades"]),
        ("Remittance inflows",
         "India is the world's largest recipient. They cushion the current account and support rural consumption - a documented predictor of Indian market returns.",
         ["Rising strongly", "Rising", "Flat", "Falling", "Falling sharply"]),
     ]),

    ("global", "Global cues & GIFT Nifty", 0.15,
     "The overnight gap. Read at 08:30, stale by 10:00.", [
        ("GIFT Nifty vs prior close",
         "The market's estimate of the opening gap, priced through the US session while NSE is shut. Compare against the prior Nifty FUTURES close, not spot. See Guide.",
         ["Above +150 pts", "+50 to +150 pts", "-50 to +50 - flat",
          "-150 to -50 pts", "Below -150 pts"]),
        ("US close (S&P 500 / Nasdaq)",
         "Sets the risk tone. Nasdaq matters more than the Dow for Indian IT and for broad risk appetite.",
         ["Up more than 1%", "Up 0.3 to 1%", "Flat", "Down 0.3 to 1%", "Down more than 1%"]),
        ("Asian markets this morning",
         "Nikkei, Hang Seng and Kospi trade alongside our session and often confirm or contradict GIFT Nifty in real time.",
         ["Broadly strong", "Mildly positive", "Mixed", "Mildly negative", "Broadly weak"]),
        ("Dollar index (DXY)",
         "A rising dollar pulls capital out of emerging markets - one of the more reliable medium-term headwinds for FII flows.",
         ["Falling sharply", "Falling", "Flat", "Rising", "Rising sharply"]),
        ("US 10-year Treasury yield",
         "Higher risk-free yields raise the hurdle for emerging-market equity and compress the multiple India trades at.",
         ["Falling sharply", "Falling", "Flat", "Rising", "Rising sharply"]),
        ("US VIX",
         "Global risk appetite. Above 25 tends to drag every emerging market down regardless of local fundamentals.",
         ["Below 14 - calm", "14 to 17", "17 to 20", "20 to 25 - elevated", "Above 25 - global fear"]),
     ]),

    ("deriv", "Derivatives & positioning", 0.30,
     "Where the money actually sits. Turns weekly, around expiry.", [
        ("India VIX level",
         "Annualised expected move from the Nifty option order book. Not directional - but the regime it describes decides position size. See Guide.",
         ["12 to 15 - calm, trend-friendly", "15 to 18 - normal", "Below 12 - complacent",
          "18 to 25 - elevated", "Above 25 - fear"]),
        ("India VIX direction",
         "Rate of change beats level. VIX rising while the index also rises means protection is being bought into strength.",
         ["Falling with price rising", "Falling", "Flat", "Rising",
          "Rising with price rising - divergence"]),
        ("Index futures OI build-up",
         "Price direction paired with open interest - whether a move carries new money or is the other side unwinding. See Guide.",
         ["Long build-up - price up, OI up", "Short covering - price up, OI down",
          "No clear build-up", "Long unwinding - price down, OI down",
          "Short build-up - price down, OI up"]),
        ("Put-call ratio (OI)",
         "Put writers expecting the index to hold is bullish. Past 1.5 it inverts - positioning is crowded and the trade is full.",
         ["1.2 to 1.5 - bullish", "1.0 to 1.2 - mildly bullish", "0.8 to 1.0 - neutral",
          "Above 1.5 - crowded long, contrarian", "Below 0.8 - bearish"]),
        ("Option chain walls",
         "Highest Call OI is resistance, highest Put OI support. Watch the shift, not the level - writers rolling a strike up is the live signal.",
         ["Put writers rolling strikes up", "Put wall holding firm", "Walls unchanged",
          "Call writers adding above", "Put writers unwinding and rolling down"]),
        ("FII index futures positioning",
         "The long-short ratio. Extended readings at either end often precede a squeeze rather than a continuation.",
         ["Net long above 70%", "Net long 55 to 70%", "Balanced 45 to 55%",
          "Net short 55 to 70%", "Net short above 70%"]),
        ("FII / DII cash flows",
         "Direction and whether they agree. DII buying against FII selling slows a fall rather than reversing it.",
         ["Both net buyers", "FII buying, DII flat", "FII selling, DII absorbing",
          "Both net sellers, modest", "Both net sellers, heavy"]),
        ("Spot vs max pain",
         "Price drifts toward max pain into expiry because writers defend it. Meaningful in the last two sessions; noise before that.",
         ["Spot well below max pain", "Spot below max pain", "Spot at max pain, or early in series",
          "Spot above max pain", "Spot well above max pain"]),
     ]),

    ("tech", "Price action & technical", 0.35,
     "The trigger. Turns intraday - this is what gets you in and out.", [
        ("Market structure",
         "Higher highs with higher lows is an uptrend. The most reliable single read on a chart, and the one that needs no indicator.",
         ["Higher highs, higher lows", "Uptrend, pulling back", "Range-bound",
          "Downtrend, bouncing", "Lower highs, lower lows"]),
        ("Price vs moving averages",
         "The 20, 50 and 200 DMA. Position against the 200 DMA separates a correction within a bull market from a bear market.",
         ["Above all three, rising", "Above 200, above 50", "Mixed / tangled",
          "Below 50, above 200", "Below all three, falling"]),
        ("Break of structure",
         "A decisive break of the prior swing high or low - the cleanest evidence that control has changed hands.",
         ["Bullish BoS, held on retest", "Bullish BoS, unconfirmed", "No break",
          "Bearish BoS, unconfirmed", "Bearish BoS, held on retest"]),
        ("Candlestick signal at a level",
         "A pattern only counts where it forms. A bullish engulfing at support is information; the same candle mid-range is noise.",
         ["Bullish engulfing or pin bar at support", "Inside bar breaking up", "No clear pattern",
          "Inside bar breaking down", "Bearish engulfing or pin bar at resistance"]),
        ("Volume confirmation",
         "A breakout on thin volume is a trap in waiting. Volume separates a real break from a fakeout.",
         ["Rising volume with the move", "Volume above average", "Average volume",
          "Volume fading on the move", "Breakout on thin volume"]),
        ("RSI (14)",
         "Divergence beats level. In a strong trend RSI stays overbought for weeks - selling 70 blindly misses the best part of a move.",
         ["40 to 60 rising, no divergence", "60 to 70 in an uptrend", "Neutral, flat",
          "Above 70 with bearish divergence", "Below 30 and falling"]),
        ("MACD",
         "A lagging confirmation tool, not a trigger. Use it to validate what price action has already told you.",
         ["Bullish cross, histogram expanding", "Above signal line", "Flat around zero",
          "Below signal line", "Bearish cross, histogram expanding"]),
        ("ADX - trend strength",
         "Above 25 means a trend worth following; below 20 means range tactics. ADX says nothing about direction.",
         ["Above 30, rising with price", "25 to 30", "20 to 25", "Below 20 - choppy",
          "Above 30, rising with price falling"]),
        ("Market breadth",
         "Advance-decline and the share of stocks above their 200 DMA. An index at highs on narrowing breadth is carried by a few names.",
         ["Broad participation, A/D strong", "Healthy breadth", "Mixed",
          "Narrowing - index up, breadth down", "Broad decline"]),
     ]),
]

# Opens on a realistic case, not a tidy one: strong tape and bullish positioning
# while macro disagrees, so the confluence test has something to say on load.
EXAMPLE = {
    "macro": [2, 2, 2, 2, 3, 3, 4, 3, 2],
    "global": [1, 1, 2, 2, 3, 1],
    "deriv": [0, 1, 0, 1, 0, 1, 2, 1],
    "tech": [0, 1, 1, 0, 1, 1, 1, 1, 3],
}


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
    put(ws, "A1", title, font=Font(name=FONT, size=14, bold=True, color=WHITE),
        fill=NAVY, align="left")
    ws.row_dimensions[1].height = 30
    ws.merge_cells(f"A2:{last}2")
    put(ws, "A2", sub, font=ITAL, fill=GREY, align="left")
    ws.row_dimensions[2].height = 17


def heads(ws, row, labels, left=()):
    for i, label in enumerate(labels, start=1):
        col = get_column_letter(i)
        put(ws, f"{col}{row}", label, font=HEAD, fill=BLUE,
            align="left" if col in left else "center", wrap=True, border=True)
    ws.row_dimensions[row].height = 20


# ------------------------------------------------------------------- Key ----
def build_key(wb):
    """Option lists, five rows per indicator. Data validation and MATCH both
    point here, so the sheet is the single source for what a reading can be."""
    ws = wb.create_sheet("Key")
    ws.sheet_view.showGridLines = False
    banner(ws, "SCORING KEY",
           "Five options per indicator, always in the order +2, +1, 0, -1, -2. "
           "Do not reorder them - the Console scores a reading by its position here.", "D")
    heads(ws, 4, ["Layer", "Indicator", "Option", "Score"], left=("A", "B", "C"))
    for col, w in zip("ABCD", (22, 34, 46, 9)):
        ws.column_dimensions[col].width = w

    starts = {}
    r = 5
    for lid, lname, _w, _role, inds in LAYERS:
        for name, _help, opts in inds:
            starts[(lid, name)] = r
            for j, opt in enumerate(opts):
                put(ws, f"A{r}", lname if j == 0 else "", border=True)
                put(ws, f"B{r}", name if j == 0 else "", border=True)
                put(ws, f"C{r}", opt, border=True)
                score = 2 - j
                put(ws, f"D{r}", score, fmt=NUM0, align="center", border=True,
                    font=Font(name=FONT, size=10,
                              color=BULL if score > 0 else BEAR if score < 0 else "4C5A66"))
                r += 1
    return starts


# --------------------------------------------------------------- Console ----
def build_console(wb, starts):
    ws = wb.create_sheet("Console")
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCD", (36, 46, 10, 58)):
        ws.column_dimensions[col].width = w
    for col in ("F", "G"):
        ws.column_dimensions[col].hidden = True

    banner(ws, "NIFTY BIAS CONSOLE",
           "Set every blue cell from today's live data. Scores, layer totals and the "
           "confluence test calculate themselves.", "D")

    ws.merge_cells("A3:D3")
    put(ws, "A3", "This file holds no live market data. Every dropdown opens on a worked "
                  "example - replace each one before acting on the output. The confluence "
                  "test matters more than the headline score.",
        font=Font(name=FONT, size=9, color=FLAT), fill=FLAT_BG, wrap=True)
    ws.row_dimensions[3].height = 28

    # ---- work out where each indicator will sit before writing the summary --
    IND_START = 32
    rows = {}
    r = IND_START
    for lid, lname, _w, _role, inds in LAYERS:
        r += 1                                   # layer header row
        first = r
        for name, _h, _o in inds:
            rows[(lid, name)] = r
            r += 1
        rows[lid] = (first, r - 1)
        r += 1                                   # spacer

    # ------------------------------------------------------- verdict block --
    # The composite lives once, in hidden F6; everything on show reads from it.
    put(ws, "F6", '=IF(SUM($G$10:$G$13)=0,"",SUM($F$10:$F$13)/SUM($G$10:$G$13))',
        font=CALC, fmt='0.0')

    put(ws, "A5", "COMPOSITE BIAS", font=SECT)
    ws.merge_cells("A6:A7")
    put(ws, "A6", '=IF($F$6="","-",ROUND($F$6,0))', font=HUGE, fmt=NUM0,
        align="center", border=True)
    ws.merge_cells("B6:D6")
    put(ws, "B6", '=IF($F$6="","Nothing set yet",'
                  'IF($F$6>=50,"STRONG BULLISH",IF($F$6>=20,"BULLISH",'
                  'IF($F$6>-20,"NEUTRAL - no directional trade",'
                  'IF($F$6>-50,"BEARISH","STRONG BEARISH")))))',
        font=BIG, align="left", border=True)
    ws.merge_cells("B7:D7")
    put(ws, "B7", '=IF($F$6="","set an indicator to begin",'
                  '"composite "&TEXT($F$6,"+0.0;-0.0;0.0")&'
                  '"   -   weighted over the layers that have a view")',
        font=SMALL, align="left", border=True)
    ws.row_dimensions[6].height = 26
    ws.row_dimensions[7].height = 16

    # -------------------------------------------------------- layer summary --
    heads(ws, 9, ["Layer", "What it is", "Score", "Reading"], left=("A", "B", "D"))
    for i, (lid, lname, weight, role, _inds) in enumerate(LAYERS):
        r = 10 + i
        first, last = rows[lid]
        put(ws, f"A{r}", f"{lname}  ({weight:.0%})", font=BOLD, border=True)
        put(ws, f"B{r}", role, font=SMALL, wrap=True, border=True)
        put(ws, f"C{r}", f'=IF(COUNT($C${first}:$C${last})=0,"",'
                         f'AVERAGE($C${first}:$C${last})*50)',
            font=CALC, fmt='+0;-0;0', align="center", border=True)
        put(ws, f"D{r}", f'=IF($C{r}="","not set",'
                         f'IF($C{r}>=50,"strongly bullish",IF($C{r}>=15,"bullish",'
                         f'IF($C{r}>-15,"no clear view",IF($C{r}>-50,"bearish",'
                         f'"strongly bearish")))))',
            font=BODY, border=True)
        # hidden: weighted contribution, and the weight only when the layer has a view
        put(ws, f"F{r}", f'=IF($C{r}="",0,$C{r}*{weight})', font=CALC, fmt='0.000')
        put(ws, f"G{r}", f'=IF($C{r}="",0,{weight})', font=CALC, fmt='0.00')
        ws.row_dimensions[r].height = 26

    ws.conditional_formatting.add("C10:C13",
        DataBarRule(start_type="num", start_value=-100, end_type="num", end_value=100,
                    color="1B4D63", showValue=True))

    # ------------------------------------------------------------ confluence --
    put(ws, "A15", "CONFLUENCE TEST", font=SECT)
    # a layer only votes past +/-15, so a near-flat layer neither confirms nor denies
    up = 'COUNTIFS($C$10:$C$13,">=15")'
    dn = 'COUNTIFS($C$10:$C$13,"<=-15")'
    put(ws, "A17", "Verdict", font=BOLD, border=True)
    ws.merge_cells("B17:D17")
    put(ws, "B17", f'=IF({up}+{dn}<2,"THIN - too few layers have a view",'
                   f'IF(OR({dn}=0,{up}=0),"ALIGNED - every layer with a view agrees",'
                   f'IF(AND(MIN({up},{dn})=1,{up}+{dn}>=3),'
                   f'"MOSTLY AGREED - one layer dissents",'
                   f'"CONFLICTED - the layers disagree")))',
        font=BIG, align="left", border=True)
    put(ws, "A18", "What to do", font=BOLD, border=True)
    ws.merge_cells("B18:D18")
    put(ws, "B18", f'=IF({up}+{dn}<2,"Fill in more indicators before reading the score.",'
                   f'IF(OR({dn}=0,{up}=0),'
                   f'"This is the setup the framework exists to find. The composite can be acted on.",'
                   f'IF(AND(MIN({up},{dn})=1,{up}+{dn}>=3),'
                   f'"Workable at reduced size. The dissenting layer is what to watch for the trade going wrong.",'
                   f'"The composite is averaging two opposing views, not signalling. Stand aside.")))',
        font=BODY, wrap=True, border=True)
    ws.row_dimensions[18].height = 28
    put(ws, "A19", "Layers voting", font=BOLD, border=True)
    ws.merge_cells("B19:D19")
    put(ws, "B19", f'={up}&" up / "&{dn}&" down  (a layer votes only past +/-15; '
                   f'the rest are too flat to count)"',
        font=SMALL, align="left", border=True)

    # ------------------------------------------------------------ the chart --
    chart = BarChart()
    chart.type = "col"
    chart.title = "Layer scores"
    chart.y_axis.title = "bias  (-100 to +100)"
    chart.y_axis.scaling.min = -100
    chart.y_axis.scaling.max = 100
    chart.height, chart.width = 8.4, 17
    chart.add_data(Reference(ws, min_col=3, min_row=9, max_row=13), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=10, max_row=13))
    chart.legend = None
    fix_chart(chart, cat_ref="'Console'!$A$10:$A$13")
    ws.add_chart(chart, "A21")

    # -------------------------------------------------------- the indicators --
    r = IND_START
    for lid, lname, weight, role, inds in LAYERS:
        ws.merge_cells(f"A{r}:D{r}")
        put(ws, f"A{r}", f"{lname.upper()}   -   {weight:.0%} of the composite",
            font=HEAD, fill=BLUE, border=True)
        ws.row_dimensions[r].height = 20
        r += 1
        for name, helptext, opts in inds:
            k = starts[(lid, name)]
            rng = f"Key!$C${k}:$C${k + 4}"
            put(ws, f"A{r}", name, font=BOLD, border=True)
            put(ws, f"B{r}", None, font=PICK, fill=SUNK, border=True)
            put(ws, f"C{r}", f'=IF($B{r}="","",3-MATCH($B{r},{rng},0))',
                font=CALC, fmt=NUM0, align="center", border=True)
            put(ws, f"D{r}", helptext, font=SMALL, wrap=True, border=True)
            ws.row_dimensions[r].height = 30
            dv = DataValidation(type="list", formula1=f"={rng}", allow_blank=True,
                                showDropDown=False, errorTitle="Not one of the options",
                                error="Pick a reading from the list, or clear the cell.")
            ws.add_data_validation(dv)
            dv.add(f"B{r}")
            r += 1
        r += 1

    # worked example, so the sheet opens in a working state
    for lid, lname, _w, _role, inds in LAYERS:
        for idx, (name, _h, opts) in enumerate(inds):
            ws[f"B{rows[(lid, name)]}"] = opts[EXAMPLE[lid][idx]]

    ws.freeze_panes = "A5"
    return ws


# ----------------------------------------------------------------- Guide ----
GIFT = [
    ("> +150 pts", "Strong gap up",
     "Gap-and-go, or an exhaustion gap. Wait for the first 15-minute candle to close before committing."),
    ("+50 to +150", "Gap up",
     "Healthy. Needs the opening range high to break for continuation."),
    ("-50 to +50", "Flat",
     "No overnight information. Trade the intraday structure, not the gap."),
    ("-150 to -50", "Gap down",
     "Check whether it fills. An unfilled gap down by noon signals genuine supply."),
    ("< -150 pts", "Strong gap down",
     "Panic opens often mark intraday lows. Let the first 30 minutes print before shorting."),
]
VIX = [
    ("Below 12", "Complacency",
     "Options are cheap. Good for buyers, dangerous for naked sellers - a shock is underpriced."),
    ("12 - 15", "Calm", "The trend-following regime. Breakouts tend to follow through."),
    ("15 - 20", "Normal", "Nothing unusual. Trade the structure as it is."),
    ("20 - 25", "Elevated",
     "Widen stops or cut size - the same stop gets hit by noise. Premium selling pays better."),
    ("Above 25", "Fear",
     "Historically closer to bottoms than tops. Contrarian longs work; leverage does not."),
]
OI = [
    ("Up", "Up", "Long build-up - new buying, the strongest bullish state"),
    ("Down", "Up", "Short build-up - new selling, the strongest bearish state"),
    ("Up", "Down", "Short covering - a rally on exits, not conviction; it fades"),
    ("Down", "Down", "Long unwinding - holders leaving, weak hands first"),
]


def build_guide(wb):
    ws = wb.create_sheet("Guide")
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCD", (20, 22, 74, 10)):
        ws.column_dimensions[col].width = w
    banner(ws, "INDICATOR GUIDE", "How to read the inputs the Console scores.", "C")

    r = 4

    def para(text, rows=2, font=BODY, fill=None):
        nonlocal r
        ws.merge_cells(f"A{r}:C{r + rows - 1}")
        put(ws, f"A{r}", text, font=font, fill=fill, wrap=True)
        for k in range(rows):
            ws.row_dimensions[r + k].height = 15
        r += rows + 1

    def table(title, cols, body):
        nonlocal r
        put(ws, f"A{r}", title, font=SECT)
        r += 1
        for i, c in enumerate(cols):
            put(ws, f"{get_column_letter(1 + i)}{r}", c, font=HEAD, fill=BLUE,
                align="left", border=True)
        r += 1
        for row in body:
            for i, v in enumerate(row):
                put(ws, f"{get_column_letter(1 + i)}{r}", v,
                    font=BOLD if i < 2 else BODY, wrap=i == 2, border=True)
            ws.row_dimensions[r].height = 26
            r += 1
        r += 1

    put(ws, f"A{r}", "GIFT NIFTY", font=SECT); r += 1
    para("GIFT Nifty is the Nifty 50 futures contract traded on NSE International Exchange "
         "in GIFT City, Gandhinagar. It took over the offshore Nifty contract from SGX Nifty "
         "in Singapore in July 2023 and settles in US dollars. It runs roughly 21 hours a day "
         "across two sessions - about 06:30 to 15:40 IST, then 16:35 to 02:45 IST - so it "
         "stays open through the US cash session while the Indian market is shut.", 3)
    para("That overlap is what makes it the most useful pre-open input. Anything that happens "
         "overnight is absorbed by GIFT Nifty before NSE opens. Take its level against the "
         "previous Nifty FUTURES close, not the spot close, since spot still carries the cost "
         "of carry. The difference is the market's estimate of the opening gap.", 3)
    para("Two cautions. GIFT Nifty indicates the open, not the day: a strong gap up sold into "
         "within the first fifteen minutes is a classic distribution signal and often marks the "
         "day's high. And its volumes are far thinner than NSE cash, so a large late move on "
         "light volume can misprice the gap.", 3)
    table("Reading the gap", ["Gap vs prior close", "Reading", "What to watch at 09:15"], GIFT)

    put(ws, f"A{r}", "INDIA VIX", font=SECT); r += 1
    para("India VIX is computed by NSE from the order book of near- and next-month Nifty 50 "
         "options, on the methodology CBOE developed for the US VIX. It reads as an annualised "
         "percentage: a VIX of 14 means the option market is pricing a 14% move over the next "
         "year, which is roughly 14 / SQRT(12) = 4.0% over the next month, one standard "
         "deviation. Divide by SQRT(252) for a daily figure - about 0.88% at VIX 14.", 3)
    para("It is not directional. It measures expected magnitude, not sign. But because fear is "
         "priced faster than greed it carries an inverse correlation with Nifty of roughly "
         "-0.7 to -0.8. The asymmetry matters more than the correlation: VIX spikes mark "
         "bottoms far more reliably than low VIX marks tops. A market can sit at VIX 11 for "
         "months while grinding higher. It cannot sit at VIX 35 for long.", 3)
    table("Regimes", ["India VIX", "Regime", "What it means for positioning"], VIX)
    para("Read the rate of change alongside the level. VIX jumping from 13 to 17 in a session "
         "is a bigger signal than VIX sitting at 19. And watch the divergence that precedes "
         "trouble: Nifty making a new high while VIX also rises means the option market is "
         "paying up for protection into strength, which rarely resolves upward.", 3)

    put(ws, f"A{r}", "OPEN INTEREST, PCR AND MAX PAIN", font=SECT); r += 1
    para("Open interest counts contracts outstanding. Paired with price direction it tells you "
         "whether a move is backed by new money or is merely being unwound.", 2)
    table("OI build-up", ["Price", "Open interest", "Reading"], OI)
    para("Put-call ratio is total Put OI divided by Call OI. Because option writers are the "
         "better-informed side, heavy put writing means writers expect the index to hold - "
         "bullish, the opposite of how it first reads. Below 0.8 bearish, 0.8 to 1.2 neutral, "
         "above 1.2 bullish. Past 1.5 or below 0.6 it inverts into a contrarian extreme: "
         "everyone is already positioned.", 3)
    para("Max pain is the strike at which the largest rupee value of options expires worthless. "
         "Price drifts toward it into expiry because writers defend it. Useful on expiry day and "
         "the day before; noise at the start of a series.", 2)
    para("The walls: highest Call OI acts as resistance, highest Put OI as support. The shift "
         "matters more than the level - put writers rolling a strike up is a live bullish "
         "signal; a wall sitting still is not.", 2)

    put(ws, f"A{r}", "THE MODEL LAYER", font=SECT); r += 1
    para("BiLSTM runs a long short-term memory network over the sequence in both directions, so "
         "each point is encoded with the context that follows it as well as what preceded it. "
         "It is the standard strong baseline for financial time series. The bidirectional pass "
         "is legitimate when training on history and becomes lookahead bias the moment it is "
         "applied to the live edge without care.", 3)
    para("Liquid neural networks are continuous-time recurrent networks whose time constants "
         "depend on the input, so the dynamics keep adapting after training rather than being "
         "frozen. They need far fewer neurons than an LSTM and hold up better on irregular, "
         "non-stationary series - which is the honest description of an index. That adaptivity "
         "is why they are interesting for regime shifts.", 3)
    para("Hybrid ensembles pair a deep sequence model with gradient-boosted trees on engineered "
         "features and a sentiment channel over financial news. For India the multilingual part "
         "is not decoration: material news breaks in Hindi, Gujarati and Marathi business media "
         "before it reaches English wires.", 3)
    para("What the published accuracy figures mean: almost all are directional accuracy on a "
         "held-out slice of history, not live profit. Four things separate the two - "
         "non-stationarity, since a model fitted before 2020 did not survive 2020; transaction "
         "costs and slippage, which consume most small edges; survivorship and lookahead bias in "
         "the dataset; and adaptivity, because a published edge is one being arbitraged away. "
         "The defensible use is probability and regime classification, never a point forecast of "
         "a closing level.", 4)

    put(ws, f"A{r}", "USING IT", font=SECT); r += 1
    for line in [
        "Agreement is the signal. A composite of +45 with all four layers pointing the same way "
        "is a trade. The same +45 built from strong technicals fighting bearish derivatives is "
        "two crowds disagreeing, and the Console flags it.",
        "Layers move on different clocks. Macro sets the regime over quarters. Derivatives turn "
        "weekly, around expiry. Price action turns intraday. Do not re-read macro every morning; "
        "do re-read the option chain.",
        "Price action is the trigger, not the thesis. Derivative data shows where positioning is "
        "crowded; price action shows when it breaks. Entering on positioning alone means being "
        "early, which in a leveraged account is indistinguishable from being wrong.",
        "Size against VIX, not against conviction. The stop that is right at VIX 12 gets taken "
        "out by noise at VIX 24. Same conviction, smaller position.",
        "A neutral reading is a position. A composite between -20 and +20 means the evidence "
        "does not support a directional trade. Standing aside is the correct output.",
    ]:
        para(line, 2)

    ws.merge_cells(f"A{r}:C{r + 2}")
    put(ws, f"A{r}", "For educational and analytical use. Nothing here is investment advice or a "
                     "recommendation to buy, sell or hold any security, index or derivative. "
                     "Derivatives carry a risk of loss exceeding the capital deployed. Indicator "
                     "thresholds are conventions in common use, not settled facts. Verify every "
                     "input against the primary source - NSE for option chain, OI, VIX and FII or "
                     "DII activity; NSE International Exchange for GIFT Nifty; MoSPI and RBI for "
                     "macro releases - before acting.",
        font=ITAL, wrap=True)
    return ws


def main(out="Nifty_Bias_Console.xlsx"):
    wb = Workbook()
    wb.remove(wb.active)
    starts = build_key(wb)
    build_console(wb, starts)
    build_guide(wb)
    wb.move_sheet("Console", offset=-2)
    wb.active = 0
    wb.save(out)
    print(f"wrote {out}")


if __name__ == "__main__":
    main(*(sys.argv[1:2] or []))
