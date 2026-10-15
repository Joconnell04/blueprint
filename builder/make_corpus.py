#!/usr/bin/env python3
"""make_corpus.py: write the import-scan test corpus into builder/corpus/.

Usage: python3 builder/make_corpus.py [out_dir]

Seven workbooks from docs/research/import-scan.md section 5 (file 1, the house
model, is builder/sample_house.xlsx built by build_xlsx.py) plus expected.json
with the blocks the scanner should find: sheet, type, title cell, extent, and
optional field checks. Also writes style_template.xlsx, an example template for
build_xlsx.py's template mode with a "Style" sheet (not scanned, not in
expected.json). Python 3.9, openpyxl only.
"""

import json
import os
import random
import sys

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table

HERE = os.path.dirname(os.path.abspath(__file__))
BOLD = Font(bold=True)
BLUE = Font(color="0000FF")
NOTE = Font(italic=True, color="595959")
THIN = Side(style="thin")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def exp(sheet, btype, row, col, rows, cols, min_conf=0.0, **fields):
    d = {"sheet": sheet, "type": btype, "row": row, "col": col, "rows": rows, "cols": cols,
         "min_conf": min_conf}
    d.update(fields)
    return d


def pairs(ws, row, col, items, title=None, blue=True, note=None):
    """Title (bold) then label/value pairs. Returns the number of rows written."""
    if title:
        ws.cell(row, col, title).font = BOLD
        if note:
            ws.cell(row, col + 2, note).font = NOTE
        row += 1
    for label, value in items:
        ws.cell(row, col, label)
        c = ws.cell(row, col + 1, value)
        if blue and not (isinstance(value, str) and value.startswith("=")):
            c.font = BLUE
        row += 1
    return row


def table(ws, row, col, headers, data, title=None, total=None, bordered=False, note=None):
    if title:
        ws.cell(row, col, title).font = BOLD
        if note:
            ws.cell(row, col + len(headers), note).font = NOTE
        row += 1
    for j, h in enumerate(headers):
        c = ws.cell(row, col + j, h)
        c.font = BOLD
        if bordered:
            c.border = BOX
    row += 1
    for data_row in data:
        for j, v in enumerate(data_row):
            c = ws.cell(row, col + j, v)
            if bordered:
                c.border = BOX
        row += 1
    if total:
        for j, v in enumerate(total):
            c = ws.cell(row, col + j, v)
            c.font = BOLD
            if bordered:
                c.border = BOX
        row += 1
    return row


# ------------------------------------------------------------- workbooks


def plain_financial(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Model"
    ws.cell(1, 1, "Assumptions")
    items = [("Growth", 0.1), ("COGS share", 0.4), ("S&M share", 0.12), ("R&D share", 0.08),
             ("G&A share", 0.06), ("Tax rate", 0.21), ("Interest ($K)", 40)]
    pairs(ws, 2, 1, items, blue=True)
    ws.cell(10, 1, "P&L")
    years = ["2024", "2025", "2026", "2027", "2028", "2029"]
    ws.cell(11, 1, "Line item")
    for j, y in enumerate(years):
        c = ws.cell(11, 2 + j, y)
        c.number_format = "@"
    lines = [("Revenue", "rev"), ("COGS", "=%s12*$B$3"), ("Gross profit", "=%s12-%s13"),
             ("Sales & marketing", "=%s12*$B$4"), ("R&D", "=%s12*$B$5"), ("G&A", "=%s12*$B$6"),
             ("Hosting", "=%s12*0.03"), ("Support", "=%s12*0.02"), ("Payroll", "=%s12*0.1"),
             ("Rent", "=%s12*0.01"), ("Travel", "=%s12*0.005"), ("Other", "=%s12*0.004"),
             ("EBITDA", "=%s14-SUM(%s15:%s22)"), ("D&A", "da"), ("Interest", "=$B$8"),
             ("EBIT", "=%s23-%s24-%s25"), ("Tax", "=%s26*$B$7"), ("Net income", "=%s26-%s27")]
    for i, (label, f) in enumerate(lines):
        r = 12 + i
        cell = ws.cell(r, 1, label)
        if label in ("Gross profit", "EBITDA", "Net income"):
            cell.font = BOLD
        for j in range(6):
            L = get_column_letter(2 + j)
            if f == "rev":
                v = 1000 if j == 0 else "=%s12*(1+$B$2)" % get_column_letter(1 + j)
            elif f == "da":
                v = 50
            else:
                v = f.replace("%s", L)
            c = ws.cell(r, 2 + j, v)
            if label in ("Gross profit", "EBITDA", "Net income"):
                c.font = BOLD
    ws.freeze_panes = "B12"
    wb.save(path)
    return [exp("Model", "inputs", 1, 1, 8, 2, 0.7),
            exp("Model", "calc", 10, 1, 20, 7, 0.75, subtotals=2, has_total=True)], 0


def data_dump(path):
    rnd = random.Random(3)
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    headers = ["id", "region", "status"] + ["m%d" % i for i in range(1, 12)]
    for j, h in enumerate(headers):
        ws.cell(1, 1 + j, h)
    regions = ["North", "South", "East", "West"]
    for r in range(2, 3001):
        ws.cell(r, 1, "row-%05d" % r)
        ws.cell(r, 2, rnd.choice(regions))
        ws.cell(r, 3, rnd.choice(["open", "closed"]))
        for j in range(11):
            ws.cell(r, 4 + j, round(rnd.random() * 1000, 2))
    ws.auto_filter.ref = "A1:N3000"
    ws.column_dimensions["H"].hidden = True
    wb.save(path)
    return [exp("Data", "table", 1, 1, 3000, 14, 0.6, max_conf=0.8, dataRows=2999, hidden_cols=1,
                sample_len=200)], 0


def side_by_side(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Flow"
    expected = []
    for top in (2, 13):
        pairs(ws, top, 2, [("Rate", 0.05), ("Volume", 120), ("Price", 9.5), ("Churn", 0.1),
                           ("Seats", 4), ("Term", 12)], title="Inputs", note="Refresh quarterly")
        expected.append(exp("Flow", "inputs", top, 2, 7, 2, 0.7, note="Refresh quarterly"))
        data = [["Row %d" % i, i * 10, i * 3, i * 7, i * 2] for i in range(1, 8)]
        table(ws, top, 6, ["Item", "A", "B", "C", "D"], data, title="Table", note="Source: CRM")
        expected.append(exp("Flow", "table", top, 6, 9, 5, 0.7, note="Source: CRM"))
        pairs(ws, top, 13, [("Sum A", "=SUM(G%d:G%d)" % (top + 2, top + 8)),
                            ("Sum B", "=SUM(H%d:H%d)" % (top + 2, top + 8)),
                            ("Max C", "=MAX(I%d:I%d)" % (top + 2, top + 8)),
                            ("Avg D", "=AVERAGE(J%d:J%d)" % (top + 2, top + 8))], title="Outputs")
        expected.append(exp("Flow", "output", top, 13, 5, 2, 0.7))
    wb.save(path)
    return expected, 0


def chart_heavy(path):
    wb = Workbook()
    data = wb.active
    data.title = "Data"
    rows = [["Jan", 100, 60, "=C4-D4", 10], ["Feb", 120, 70, "=C5-D5", 12], ["Mar", 140, 75, "=C6-D6", 14],
            ["Apr", 130, 72, "=C7-D7", 13], ["May", 150, 80, "=C8-D8", 15]]
    table(data, 2, 2, ["Month", "Sales", "Costs", "Margin", "Units"], rows, title="Monthly", bordered=True)
    charts = wb.create_sheet("Charts")
    crow = [["North", 10, 12, 14], ["South", 8, 9, 11], ["East", 6, 7, 9], ["West", 5, 6, 8], ["Central", 4, 5, 6]]
    table(charts, 2, 2, ["Region", "Q1", "Q2", "Q3"], crow, title="Regions", bordered=True)
    expected = [exp("Data", "table", 2, 2, 7, 5, 0.7), exp("Charts", "table", 2, 2, 7, 4, 0.7)]

    def series(ws_name, chart, cols, cats_col=2, r1=3, r2=8):
        ws = wb[ws_name]
        for c in cols:
            chart.add_data(Reference(ws, min_col=c, min_row=r1, max_row=r2), titles_from_data=True)
        chart.set_categories(Reference(ws, min_col=cats_col, min_row=r1 + 1, max_row=r2))

    bar = BarChart()
    bar.title = "Q1 by region"
    series("Charts", bar, [3])
    bar.width, bar.height = 8.5, 7.5
    charts.add_chart(bar, "G2")
    expected.append(exp("Charts", "chart", 2, 7, 15, 5, 0.9, chart_kind="bar", source_type="table"))

    line = LineChart()
    line.title = "Quarters"
    series("Charts", line, [3, 4, 5])
    anchor = TwoCellAnchor(_from=AnchorMarker(col=6, row=19), to=AnchorMarker(col=10, row=33))
    charts.add_chart(line, anchor)
    expected.append(exp("Charts", "chart", 20, 7, 15, 5, 0.9, chart_kind="line", series_cols=[1, 2, 3]))

    pie = PieChart()
    pie.title = "Q3 share"
    series("Charts", pie, [5])
    pie.width, pie.height = 8.5, 7.5
    charts.add_chart(pie, "G38")
    expected.append(exp("Charts", "chart", 38, 7, 15, 5, 0.9, chart_kind="pie"))

    combo = BarChart()
    combo.title = "Combo"
    series("Charts", combo, [3])
    combo_line = LineChart()
    series("Charts", combo_line, [4])
    combo += combo_line
    combo.width, combo.height = 8.5, 7.5
    charts.add_chart(combo, "G56")
    expected.append(exp("Charts", "chart", 56, 7, 15, 5, 0.9, chart_kind="bar"))

    cross = BarChart()
    cross.title = "Sales from Data"
    series("Data", cross, [3], cats_col=2, r1=3, r2=8)
    cross.width, cross.height = 8.5, 7.5
    charts.add_chart(cross, "G74")
    expected.append(exp("Charts", "chart", 74, 7, 15, 5, 0.7, max_conf=0.7, chart_kind="bar",
                        source_sheet="Data"))
    wb.save(path)
    return expected, 0


def messy(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Messy"
    ws.cell(1, 1, "See the review email")  # comment-like, under 40 chars
    ws.cell(100, 26, 42)
    ws.cell(3, 2, "Regional sales").font = BOLD  # far title: blank row 4 between it and the table
    table(ws, 5, 2, ["Region", "Q1", "Q2", "Q3"], [["N", 1, 2, 3], ["S", 4, 5, 6], ["E", 7, 8, 9],
                                                    ["W", 2, 3, 4], ["C", 5, 6, 7]])
    ws.cell(2, 7, "Targets").font = BOLD  # labels on the right
    for i, (v, label) in enumerate([(100, "Bookings"), (80, "Billings"), (60, "Collections"),
                                    (40, "Refunds"), (20, "Credits")]):
        ws.cell(3 + i, 7, v)
        ws.cell(3 + i, 8, label)
    # table with one blank row inside its body, total at the end
    r = table(ws, 13, 2, ["Team", "2024", "2025", "2026"],
              [["Eng", 10, 12, 14], ["Sales", 8, 9, 10], ["Ops", 4, 5, 6], ["Support", 3, 4, 5]],
              title="Opex by team")
    r += 1  # blank spacer row 19
    for i, row in enumerate([["Finance", 2, 2, 3], ["Legal", 1, 1, 2], ["HR", 2, 2, 2], ["Exec", 3, 3, 3]]):
        for j, v in enumerate(row):
            ws.cell(r + i, 2 + j, v)
    r += 4
    for j, v in enumerate(["Total", "=SUM(C15:C23)", "=SUM(D15:D23)", "=SUM(E15:E23)"]):
        ws.cell(r, 2 + j, v).font = BOLD
    # two tables stacked with no blank row: total row directly followed by a bold title
    r = table(ws, 27, 2, ["Cost", "2024", "2025", "2026"],
              [["Cloud", 5, 6, 7], ["Tools", 1, 1, 2], ["Travel", 2, 2, 2], ["Office", 3, 3, 3]],
              title="Costs", total=["Total", "=SUM(C29:C32)", "=SUM(D29:D32)", "=SUM(E29:E32)"])
    table(ws, r, 2, ["Product", "2024", "2025", "2026"],
          [["Basic", 10, 11, 12], ["Pro", 20, 22, 24], ["Team", 30, 33, 36], ["Max", 40, 44, 48]],
          title="Prices")
    # merged two-level header
    ws.cell(42, 2, "Headcount").font = BOLD
    ws.cell(43, 3, "Plan").font = BOLD
    ws.cell(43, 3).alignment = Alignment(horizontal="center")
    ws.merge_cells("C43:E43")
    for j, h in enumerate(["Team", "2024", "2025", "2026"]):
        ws.cell(44, 2 + j, h).font = BOLD
    for i, row in enumerate([["Eng", 10, 12, 14], ["Sales", 8, 9, 10], ["Ops", 4, 5, 6], ["Exec", 3, 3, 3]]):
        for j, v in enumerate(row):
            ws.cell(45 + i, 2 + j, v)
    wb.save(path)
    expected = [
        exp("Messy", "table", 4, 2, 7, 4, 0.5),                       # far title left behind, no title row
        exp("Messy", "inputs", 2, 7, 6, 2, 0.5, labels="right"),
        exp("Messy", "table", 13, 2, 12, 4, 0.5, max_conf=0.75, bridged=True, has_total=True),
        exp("Messy", "table", 27, 2, 7, 4, 0.6, has_total=True),
        exp("Messy", "table", 34, 2, 6, 4, 0.6),
        exp("Messy", "table", 42, 2, 7, 4, 0.5, header_rows=2),
    ]
    return expected, 3  # A1 lone-text, B3 lone-text, Z100 stray-cell


def listobject(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    ws.cell(2, 2, "Sales table").font = BOLD
    headers = ["Rep", "Region", "Units", "Price", "Amount"]
    for j, h in enumerate(headers):
        ws.cell(3, 2 + j, h)
    rows = [["Ann", "N", 10, 5, "=D4*E4"], ["Bob", "S", 8, 6, "=D5*E5"], ["Cy", "E", 12, 4, "=D6*E6"],
            ["Di", "W", 7, 9, "=D7*E7"], ["Ed", "N", 9, 5, "=D8*E8"]]
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            ws.cell(4 + i, 2 + j, v)
    ws.cell(9, 2, "Total")
    ws.cell(9, 4, "=SUBTOTAL(109,D4:D8)")
    ws.cell(9, 6, "=SUBTOTAL(109,F4:F8)")
    tab = Table(displayName="SalesTbl", ref="B3:F9")
    tab.totalsRowCount = 1
    ws.add_table(tab)
    pairs(ws, 12, 2, [("Total amount", "=SUM(SalesTbl[Amount])"), ("Total units", "=SUM(SalesTbl[Units])"),
                      ("Average price", "=AVERAGE(SalesTbl[Price])")], title="Summary", blue=False)
    wb.save(path)
    return [exp("Sales", "table", 2, 2, 8, 5, 0.9, has_total=True, listobject=True),
            exp("Sales", "output", 12, 2, 4, 2, 0.7)], 0


def hidden_frozen(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Grouped"
    ws.cell(1, 3, "Plan (FTE)").font = BOLD
    ws.cell(1, 3).alignment = Alignment(horizontal="center")
    ws.merge_cells("C1:H1")
    for j, h in enumerate(["Team", "2024", "2025", "helper", "2026", "2027", "2028"]):
        ws.cell(2, 2 + j, h).font = BOLD
    teams = ["Eng", "Sales", "Ops", "Support", "Finance", "Legal", "HR", "Exec", "Data", "Design", "QA", "IT"]
    for i, t in enumerate(teams):
        ws.cell(3 + i, 2, t)
        for j in range(6):
            ws.cell(3 + i, 3 + j, (i + 1) * (j + 2))
    for r in (6, 9, 12):
        ws.row_dimensions[r].hidden = True
        ws.row_dimensions[r].outlineLevel = 1
    ws.column_dimensions["E"].hidden = True
    ws.column_dimensions["E"].outlineLevel = 1
    ws.freeze_panes = "C3"
    wb.save(path)
    return [exp("Grouped", "table", 1, 2, 14, 7, 0.5, max_conf=0.65, header_rows=2, hidden_rows=3,
                hidden_cols=1)], 0


def style_template(path):
    """A template workbook for build_xlsx.py's template mode.

    The "Style" sheet carries one example cell per role in column B, named by
    the label in column A (Title, Header, Label, Input, Formula, Total, Note,
    Text); the builder clones font, fill, border, alignment and number format
    from each, and takes the label, value and note column widths from columns
    B, C and D. The "Sample" sheet just shows the base font in context.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Style"
    navy = "1F3864"
    steel = "44546A"
    thin_grey = Side(style="thin", color="BFBFBF")
    box_grey = Border(left=thin_grey, right=thin_grey, top=thin_grey, bottom=thin_grey)
    rows = [
        ("Title", "Key Assumptions",
         Font(name="Georgia", size=13, bold=True, color=navy),
         PatternFill(), Border(bottom=Side(style="medium", color=navy)),
         Alignment(horizontal="center", vertical="center"), "General"),
        ("Header", "Year 1",
         Font(name="Georgia", size=11, bold=True, color="FFFFFF"),
         PatternFill(fill_type="solid", fgColor=steel), box_grey,
         Alignment(horizontal="center", vertical="center", wrap_text=True), "General"),
        ("Label", "Market size ($M)",
         Font(name="Georgia", size=11), PatternFill(), box_grey,
         Alignment(horizontal="left", vertical="center"), "General"),
        ("Input", 1200,
         Font(name="Georgia", size=11, color="1F5F99"),
         PatternFill(fill_type="solid", fgColor="FFF9E5"), box_grey,
         Alignment(horizontal="right", vertical="center"), "General"),
        ("Formula", "=B5*2",
         Font(name="Georgia", size=11, color="000000"), PatternFill(), box_grey,
         Alignment(horizontal="right", vertical="center"), "General"),
        ("Total", 2400,
         Font(name="Georgia", size=11, bold=True),
         PatternFill(fill_type="solid", fgColor="EDEDED"),
         Border(left=thin_grey, right=thin_grey, top=Side(style="thin", color=navy),
                bottom=Side(style="double", color=navy)),
         Alignment(horizontal="right", vertical="center"), "General"),
        ("Note", "Analyst estimate; see A12",
         Font(name="Georgia", size=10, italic=True, color="7F7F7F"), PatternFill(), Border(),
         Alignment(horizontal="left", vertical="top", wrap_text=False), "General"),
        ("Text", "Everything on this tab feeds the model.",
         Font(name="Georgia", size=11), PatternFill(), Border(),
         Alignment(horizontal="left", vertical="top", wrap_text=True), "General"),
    ]
    ws.cell(1, 1, "Role").font = BOLD
    ws.cell(1, 2, "Example").font = BOLD
    ws.cell(1, 3, "What the builder clones").font = BOLD
    for i, (role, example, font, fill, border, align, nf) in enumerate(rows):
        r = 2 + i
        ws.cell(r, 1, role)
        c = ws.cell(r, 2, example)
        c.font, c.fill, c.border, c.alignment, c.number_format = font, fill, border, align, nf
        ws.cell(r, 3, "font, fill, border, alignment, number format").font = NOTE
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 28   # label column width
    ws.column_dimensions["C"].width = 14   # value column width
    ws.column_dimensions["D"].width = 44   # note column width
    sample = wb.create_sheet("Sample")
    sample.cell(1, 1, "Body copy in the template's base font.").font = Font(name="Georgia", size=11)
    sample.cell(2, 1, "A bold heading").font = Font(name="Georgia", size=13, bold=True, color=navy)
    wb.save(path)


FILES = [
    ("plain_financial.xlsx", plain_financial),
    ("data_dump.xlsx", data_dump),
    ("side_by_side.xlsx", side_by_side),
    ("chart_heavy.xlsx", chart_heavy),
    ("messy.xlsx", messy),
    ("listobject.xlsx", listobject),
    ("hidden_frozen.xlsx", hidden_frozen),
]

# Written next to FILES but not scanned for expected.json: template-mode examples.
EXTRA_FILES = [
    ("style_template.xlsx", style_template),
]


def make(out_dir):
    os.makedirs(out_dir, exist_ok=True)
    expected = {}
    for name, fn in FILES:
        blocks, unclassified = fn(os.path.join(out_dir, name))
        expected[name] = {"blocks": blocks, "unclassified": unclassified}
    with open(os.path.join(out_dir, "expected.json"), "w") as fh:
        json.dump(expected, fh, indent=2)
    for name, fn in EXTRA_FILES:
        fn(os.path.join(out_dir, name))
    return expected


if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "corpus")
    result = make(target)
    print(json.dumps({"ok": True, "files": len(result), "dir": os.path.abspath(target)}))
