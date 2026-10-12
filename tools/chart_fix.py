"""Make an openpyxl chart render in Excel.

openpyxl's chart defaults are wrong in three ways that Excel will not forgive,
while LibreOffice and most previewers render the chart anyway - so a broken
chart looks fine everywhere except the place it matters:

  1. Both axes default to axPos "l". On a column chart the category axis has to
     sit at the bottom; two axes claiming the left is malformed.
  2. set_categories() always writes a numRef, even when the category cells hold
     text. Excel needs a strRef for text labels.
  3. delete is left unset on both axes; Excel wants it explicitly false.

Call fix_chart(chart, cat_ref=...) after add_data and set_categories.
"""

from openpyxl.chart.data_source import AxDataSource, StrRef


def fix_chart(chart, cat_ref=None, horizontal=False):
    """Repair the axis positions and category reference of a bar or column chart.

    cat_ref  - sheet-qualified range holding the category labels, e.g.
               "'Console'!$A$10:$A$13". Pass it whenever those labels are text.
    horizontal - True for a bar (type="bar") chart, where the axes swap.
    """
    cat_pos, val_pos = ("l", "b") if horizontal else ("b", "l")
    chart.x_axis.axPos = cat_pos
    chart.y_axis.axPos = val_pos
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.x_axis.majorTickMark = "out"
    chart.y_axis.majorTickMark = "out"
    if cat_ref:
        for s in chart.series:
            s.cat = AxDataSource(strRef=StrRef(f=cat_ref))
    return chart
