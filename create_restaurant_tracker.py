#!/usr/bin/env python3
"""
Restaurant Monthly Budget & Operations Tracker - Google Sheets Creator
Uses the gws CLI to create and fully format the spreadsheet.
"""

import json
import subprocess
import sys
import os

GWS = "/root/.cargo/bin/gws"


def run_gws(args: list, input_data: str = None) -> dict:
    """Run a gws command and return parsed JSON output."""
    cmd = [GWS] + args
    result = subprocess.run(
        cmd,
        input=input_data,
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        print(f"ERROR running: {' '.join(cmd)}", file=sys.stderr)
        print(result.stderr, file=sys.stderr)
        sys.exit(1)
    try:
        return json.loads(result.stdout)
    except json.JSONDecodeError:
        print(f"Could not parse output: {result.stdout}", file=sys.stderr)
        sys.exit(1)


def create_spreadsheet(title: str) -> str:
    """Create a new spreadsheet and return its ID."""
    print(f"Creating spreadsheet: {title}")
    payload = {"properties": {"title": title}}
    result = run_gws(
        ["sheets", "spreadsheets", "create", "--json", json.dumps(payload)]
    )
    spreadsheet_id = result["spreadsheetId"]
    print(f"  Created: {result['spreadsheetUrl']}")
    return spreadsheet_id


def batch_update(spreadsheet_id: str, requests: list, description: str = ""):
    """Send a batchUpdate request."""
    if description:
        print(f"  {description} ({len(requests)} request(s))")
    payload = {"requests": requests}
    run_gws(
        [
            "sheets", "spreadsheets", "batchUpdate",
            "--params", json.dumps({"spreadsheetId": spreadsheet_id}),
            "--json", json.dumps(payload),
        ]
    )


def values_update(spreadsheet_id: str, range_: str, values: list, description: str = ""):
    """Write values to a range."""
    if description:
        print(f"  Writing: {description}")
    payload = {"values": values}
    run_gws(
        [
            "sheets", "spreadsheets", "values", "update",
            "--params", json.dumps({
                "spreadsheetId": spreadsheet_id,
                "range": range_,
                "valueInputOption": "USER_ENTERED",
            }),
            "--json", json.dumps(payload),
        ]
    )


# ── Colour helpers ────────────────────────────────────────────────────────────

def hex_to_rgb(hex_color: str) -> dict:
    """Convert #RRGGBB to {red, green, blue} fractions."""
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return {"red": r / 255, "green": g / 255, "blue": b / 255}


DARK = hex_to_rgb("#2C2C2C")
ROSE = hex_to_rgb("#E8A0A0")
WHITE = {"red": 1.0, "green": 1.0, "blue": 1.0}
LIGHT_GREY = hex_to_rgb("#F2F2F2")
FAF0F0 = hex_to_rgb("#FAF0F0")
GREEN_FILL = hex_to_rgb("#C6EFCE")
GREEN_TEXT = hex_to_rgb("#276221")
RED_FILL = hex_to_rgb("#FFC7CE")
RED_TEXT = hex_to_rgb("#9C0006")
OPEN_HEADER = hex_to_rgb("#375623")
CLOSE_HEADER = hex_to_rgb("#843C0C")
LIGHT_PINK = hex_to_rgb("#E8A0A0")


# ── Cell / Range helpers ──────────────────────────────────────────────────────

def cell_range(sheet_id: int, r1: int, c1: int, r2: int = None, c2: int = None) -> dict:
    gr = {
        "sheetId": sheet_id,
        "startRowIndex": r1,
        "startColumnIndex": c1,
    }
    if r2 is not None:
        gr["endRowIndex"] = r2
    if c2 is not None:
        gr["endColumnIndex"] = c2
    return gr


def text_fmt(bold=False, size=11, fg=None, font="Arial") -> dict:
    fmt = {"fontFamily": font, "fontSize": size, "bold": bold}
    if fg:
        fmt["foregroundColor"] = fg
    return fmt


def cell_fmt(bg=None, text=None, h_align=None, v_align=None, wrap=None, number_format=None) -> dict:
    fmt = {}
    if bg:
        fmt["backgroundColor"] = bg
    if text:
        fmt["textFormat"] = text
    if h_align:
        fmt["horizontalAlignment"] = h_align
    if v_align:
        fmt["verticalAlignment"] = v_align
    if wrap:
        fmt["wrapStrategy"] = wrap
    if number_format:
        fmt["numberFormat"] = number_format
    return fmt


def repeat_cell_req(sheet_id, r1, c1, r2, c2, cell_format: dict, fields: str) -> dict:
    return {
        "repeatCell": {
            "range": cell_range(sheet_id, r1, c1, r2, c2),
            "cell": {"userEnteredFormat": cell_format},
            "fields": fields,
        }
    }


def merge_req(sheet_id, r1, c1, r2, c2) -> dict:
    return {
        "mergeCells": {
            "range": cell_range(sheet_id, r1, c1, r2, c2),
            "mergeType": "MERGE_ALL",
        }
    }


def freeze_req(sheet_id, rows=1, cols=0) -> dict:
    return {
        "updateSheetProperties": {
            "properties": {
                "sheetId": sheet_id,
                "gridProperties": {"frozenRowCount": rows, "frozenColumnCount": cols},
            },
            "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount",
        }
    }


def tab_color_req(sheet_id, color: dict) -> dict:
    return {
        "updateSheetProperties": {
            "properties": {
                "sheetId": sheet_id,
                "tabColorStyle": {"rgbColor": color},
            },
            "fields": "tabColorStyle",
        }
    }


def col_width_req(sheet_id, col_idx: int, width_px: int) -> dict:
    return {
        "updateDimensionProperties": {
            "range": {
                "sheetId": sheet_id,
                "dimension": "COLUMNS",
                "startIndex": col_idx,
                "endIndex": col_idx + 1,
            },
            "properties": {"pixelSize": width_px},
            "fields": "pixelSize",
        }
    }


def auto_resize_req(sheet_id, start_col=0, end_col=20) -> dict:
    return {
        "autoResizeDimensions": {
            "dimensions": {
                "sheetId": sheet_id,
                "dimension": "COLUMNS",
                "startIndex": start_col,
                "endIndex": end_col,
            }
        }
    }


def border_req(sheet_id, r1, c1, r2, c2, style="SOLID", width=1) -> dict:
    b = {"style": style, "width": width, "color": {"red": 0.8, "green": 0.8, "blue": 0.8}}
    return {
        "updateBorders": {
            "range": cell_range(sheet_id, r1, c1, r2, c2),
            "top": b, "bottom": b, "left": b, "right": b,
            "innerHorizontal": b, "innerVertical": b,
        }
    }


def cond_format_req(sheet_id, r1, c1, r2, c2, formula: str, bg: dict, text_color: dict) -> dict:
    return {
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [cell_range(sheet_id, r1, c1, r2, c2)],
                "booleanRule": {
                    "condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": formula}]},
                    "format": cell_fmt(bg=bg, text=text_fmt(fg=text_color)),
                },
            },
            "index": 0,
        }
    }


# ─────────────────────────────────────────────────────────────────────────────
# SHEET BUILDERS
# ─────────────────────────────────────────────────────────────────────────────

def build_sheet1_dashboard(sid: int, reqs: list, vals: list):
    """Sheet 1 – Dashboard"""
    print("  Building: Dashboard")

    # Title row: merge A1:G1
    reqs.append(merge_req(sid, 0, 0, 1, 7))
    reqs.append(repeat_cell_req(sid, 0, 0, 1, 7,
        cell_fmt(bg=DARK, text=text_fmt(bold=True, size=18, fg=WHITE), h_align="CENTER", v_align="MIDDLE"),
        "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)"))

    # Summary card header row (row 2)
    reqs.append(repeat_cell_req(sid, 1, 0, 2, 7,
        cell_fmt(bg=ROSE, text=text_fmt(bold=True, size=12, fg=DARK), h_align="CENTER"),
        "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"))

    # Summary value rows (rows 3-4) – white bg
    reqs.append(repeat_cell_req(sid, 2, 0, 4, 7,
        cell_fmt(bg=WHITE, text=text_fmt(size=12), h_align="CENTER"),
        "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"))

    # Currency format for Revenue, Expenses, Profit cols
    for col in [1, 2, 3]:
        reqs.append(repeat_cell_req(sid, 2, col, 4, col + 1,
            cell_fmt(number_format={"type": "CURRENCY", "pattern": "$#,##0.00"}),
            "userEnteredFormat.numberFormat"))

    # Percentage format for Food Cost %, Labor Cost %
    for col in [4, 5]:
        reqs.append(repeat_cell_req(sid, 2, col, 4, col + 1,
            cell_fmt(number_format={"type": "PERCENT", "pattern": "0.00%"}),
            "userEnteredFormat.numberFormat"))

    reqs.append(freeze_req(sid, rows=0))
    reqs.append(tab_color_req(sid, DARK))
    reqs.append(auto_resize_req(sid, 0, 7))

    # Values
    vals.append(("Dashboard!A1", [
        ["RESTAURANT MONTHLY OVERVIEW"],
        ["", "Total Revenue", "Total Expenses", "Net Profit", "Food Cost %", "Labor Cost %", ""],
        ["",
         "='Income & Expenses'!D{last}".replace("{last}", "102"),
         "='Income & Expenses'!E{last}".replace("{last}", "102"),
         "='Income & Expenses'!D102-'Income & Expenses'!E102",
         "='Food Inventory & Cost'!G{last}/Dashboard!C3".replace("{last}", "52"),
         "='Labor & Schedule'!L{last}/Dashboard!C3".replace("{last}", "22"),
         ""],
    ]))


def build_sheet2_income_expenses(sid: int, reqs: list, vals: list):
    """Sheet 2 – Income & Expenses"""
    print("  Building: Income & Expenses")
    COLS = 6  # Date | Category | Description | Income | Expense | Balance

    # Header row
    reqs.append(repeat_cell_req(sid, 0, 0, 1, COLS,
        cell_fmt(bg=DARK, text=text_fmt(bold=True, size=12, fg=WHITE), h_align="CENTER"),
        "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"))

    # Data rows alternating
    for r in range(1, 101):
        bg = WHITE if r % 2 == 1 else LIGHT_GREY
        reqs.append(repeat_cell_req(sid, r, 0, r + 1, COLS,
            cell_fmt(bg=bg, text=text_fmt(size=11)),
            "userEnteredFormat(backgroundColor,textFormat)"))

    # Currency format for Income (col 3), Expense (col 4), Balance (col 5)
    for col in [3, 4, 5]:
        reqs.append(repeat_cell_req(sid, 1, col, 101, col + 1,
            cell_fmt(number_format={"type": "CURRENCY", "pattern": "$#,##0.00"}),
            "userEnteredFormat.numberFormat"))

    # Totals row (row 102, index 101)
    reqs.append(repeat_cell_req(sid, 101, 0, 102, COLS,
        cell_fmt(bg=DARK, text=text_fmt(bold=True, size=12, fg=WHITE), h_align="CENTER"),
        "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"))
    for col in [3, 4, 5]:
        reqs.append(repeat_cell_req(sid, 101, col, 102, col + 1,
            cell_fmt(number_format={"type": "CURRENCY", "pattern": "$#,##0.00"}),
            "userEnteredFormat.numberFormat"))

    # Conditional formatting: Income col green
    reqs.append(cond_format_req(sid, 1, 3, 101, 4,
        "=AND(D2<>\"\",D2>0)", GREEN_FILL, GREEN_TEXT))
    # Conditional formatting: Expense col red
    reqs.append(cond_format_req(sid, 1, 4, 101, 5,
        "=AND(E2<>\"\",E2>0)", RED_FILL, RED_TEXT))

    reqs.append(freeze_req(sid, rows=1))
    reqs.append(tab_color_req(sid, GREEN_FILL))
    reqs.append(auto_resize_req(sid, 0, COLS))

    # Sample data
    sample_rows = [
        ["Date", "Category", "Description", "Income", "Expense", "Balance"],
        ["2024-04-01", "Dine-In", "Lunch service", 1250.00, "", "=IF(D3<>\"\",D3,0)-IF(E3<>\"\",E3,0)+F2"],
        ["2024-04-01", "Food & Beverage", "Produce order", "", 320.00, "=IF(D4<>\"\",D4,0)-IF(E4<>\"\",E4,0)+F3"],
        ["2024-04-01", "Labor", "Staff wages", "", 780.00, "=IF(D5<>\"\",D5,0)-IF(E5<>\"\",E5,0)+F4"],
        ["2024-04-02", "Takeout", "Online orders", 430.00, "", "=IF(D6<>\"\",D6,0)-IF(E6<>\"\",E6,0)+F5"],
        ["2024-04-02", "Delivery", "Third-party delivery", 210.00, "", "=IF(D7<>\"\",D7,0)-IF(E7<>\"\",E7,0)+F6"],
        ["2024-04-02", "Utilities", "Electric & gas", "", 195.00, "=IF(D8<>\"\",D8,0)-IF(E8<>\"\",E8,0)+F7"],
        ["2024-04-03", "Bar Sales", "Weekend bar", 680.00, "", "=IF(D9<>\"\",D9,0)-IF(E9<>\"\",E9,0)+F8"],
        ["2024-04-03", "Rent", "Monthly rent", "", 3200.00, "=IF(D10<>\"\",D10,0)-IF(E10<>\"\",E10,0)+F9"],
        ["2024-04-03", "Marketing", "Social media ads", "", 150.00, "=IF(D11<>\"\",D11,0)-IF(E11<>\"\",E11,0)+F10"],
        ["2024-04-04", "Catering", "Private event", 1800.00, "", "=IF(D12<>\"\",D12,0)-IF(E12<>\"\",E12,0)+F11"],
    ]
    vals.append(("'Income & Expenses'!A1", sample_rows))
    # Totals row
    vals.append(("'Income & Expenses'!A102", [
        ["TOTALS", "", "", "=SUM(D2:D101)", "=SUM(E2:E101)",
         "=D102-E102"]
    ]))


def build_sheet3_inventory(sid: int, reqs: list, vals: list):
    """Sheet 3 – Food Inventory & Cost"""
    print("  Building: Food Inventory & Cost")
    COLS = 8  # Item | Category | Unit | Par Level | Current Stock | Unit Cost | Total Value | Reorder?

    # Header row
    reqs.append(repeat_cell_req(sid, 0, 0, 1, COLS,
        cell_fmt(bg=DARK, text=text_fmt(bold=True, size=12, fg=WHITE), h_align="CENTER"),
        "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"))

    # Alternating rows: white / #FAF0F0
    for r in range(1, 51):
        bg = WHITE if r % 2 == 1 else FAF0F0
        reqs.append(repeat_cell_req(sid, r, 0, r + 1, COLS,
            cell_fmt(bg=bg, text=text_fmt(size=11)),
            "userEnteredFormat(backgroundColor,textFormat)"))

    # Currency for Unit Cost (col 5) and Total Value (col 6)
    for col in [5, 6]:
        reqs.append(repeat_cell_req(sid, 1, col, 51, col + 1,
            cell_fmt(number_format={"type": "CURRENCY", "pattern": "$#,##0.00"}),
            "userEnteredFormat.numberFormat"))

    # Reorder? column conditional: REORDER in red
    reqs.append(cond_format_req(sid, 1, 7, 51, 8,
        '=G2="REORDER"', RED_FILL, RED_TEXT))
    # OK in green
    reqs.append(cond_format_req(sid, 1, 7, 51, 8,
        '=G2="OK"', GREEN_FILL, GREEN_TEXT))

    reqs.append(freeze_req(sid, rows=1))
    reqs.append(tab_color_req(sid, FAF0F0))
    reqs.append(auto_resize_req(sid, 0, COLS))

    items = [
        ["Item", "Category", "Unit", "Par Level", "Current Stock", "Unit Cost", "Total Value", "Reorder?"],
        ["Romaine Lettuce", "Produce", "Case", 3, 2, 28.00, "=E3*F3", '=IF(E3<D3,"REORDER","OK")'],
        ["Tomatoes", "Produce", "Case", 2, 3, 22.00, "=E4*F4", '=IF(E4<D4,"REORDER","OK")'],
        ["Chicken Breast", "Meat", "lb", 40, 25, 3.50, "=E5*F5", '=IF(E5<D5,"REORDER","OK")'],
        ["Beef Sirloin", "Meat", "lb", 30, 35, 8.75, "=E6*F6", '=IF(E6<D6,"REORDER","OK")'],
        ["Salmon Fillet", "Meat", "lb", 15, 10, 12.00, "=E7*F7", '=IF(E7<D7,"REORDER","OK")'],
        ["Cheddar Cheese", "Dairy", "lb", 10, 8, 4.50, "=E8*F8", '=IF(E8<D8,"REORDER","OK")'],
        ["Heavy Cream", "Dairy", "qt", 12, 14, 3.20, "=E9*F9", '=IF(E9<D9,"REORDER","OK")'],
        ["All-Purpose Flour", "Dry Goods", "lb", 25, 30, 0.65, "=E10*F10", '=IF(E10<D10,"REORDER","OK")'],
        ["Pasta (Penne)", "Dry Goods", "lb", 20, 18, 1.10, "=E11*F11", '=IF(E11<D11,"REORDER","OK")'],
        ["Olive Oil", "Dry Goods", "gal", 5, 4, 18.00, "=E12*F12", '=IF(E12<D12,"REORDER","OK")'],
        ["White Wine", "Beverages", "bottle", 12, 15, 9.00, "=E13*F13", '=IF(E13<D13,"REORDER","OK")'],
        ["Beer (Domestic)", "Beverages", "case", 8, 5, 22.00, "=E14*F14", '=IF(E14<D14,"REORDER","OK")'],
        ["Sparkling Water", "Beverages", "case", 6, 7, 14.00, "=E15*F15", '=IF(E15<D15,"REORDER","OK")'],
        ["Dish Soap", "Cleaning Supplies", "gal", 4, 3, 8.50, "=E16*F16", '=IF(E16<D16,"REORDER","OK")'],
        ["Sanitizer Solution", "Cleaning Supplies", "gal", 5, 6, 12.00, "=E17*F17", '=IF(E17<D17,"REORDER","OK")'],
    ]
    vals.append(("'Food Inventory & Cost'!A1", items))


def build_sheet4_labor(sid: int, reqs: list, vals: list):
    """Sheet 4 – Labor & Schedule"""
    print("  Building: Labor & Schedule")
    COLS = 12  # Name | Role | Mon-Sun | Total Hours | Hourly Rate | Weekly Cost

    # Header
    reqs.append(repeat_cell_req(sid, 0, 0, 1, COLS,
        cell_fmt(bg=DARK, text=text_fmt(bold=True, size=12, fg=WHITE), h_align="CENTER"),
        "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"))

    # Data rows
    for r in range(1, 21):
        bg = WHITE if r % 2 == 1 else LIGHT_GREY
        reqs.append(repeat_cell_req(sid, r, 0, r + 1, COLS,
            cell_fmt(bg=bg, text=text_fmt(size=11)),
            "userEnteredFormat(backgroundColor,textFormat)"))

    # Role column (col 1) bold
    reqs.append(repeat_cell_req(sid, 1, 1, 21, 2,
        cell_fmt(text=text_fmt(bold=True, size=11)),
        "userEnteredFormat.textFormat"))

    # Currency for Weekly Cost (col 11)
    reqs.append(repeat_cell_req(sid, 1, 11, 21, 12,
        cell_fmt(number_format={"type": "CURRENCY", "pattern": "$#,##0.00"}),
        "userEnteredFormat.numberFormat"))

    # Totals row
    reqs.append(repeat_cell_req(sid, 21, 0, 22, COLS,
        cell_fmt(bg=DARK, text=text_fmt(bold=True, size=12, fg=WHITE), h_align="CENTER"),
        "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"))
    reqs.append(repeat_cell_req(sid, 21, 11, 22, 12,
        cell_fmt(number_format={"type": "CURRENCY", "pattern": "$#,##0.00"}),
        "userEnteredFormat.numberFormat"))

    reqs.append(freeze_req(sid, rows=1))
    reqs.append(tab_color_req(sid, hex_to_rgb("#4472C4")))
    reqs.append(auto_resize_req(sid, 0, COLS))

    employees = [
        ["Employee Name", "Role", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun",
         "Total Hours", "Hourly Rate", "Weekly Cost"],
        ["Maria Santos", "Manager", 8, 8, 8, 8, 8, 0, 0, "=SUM(C3:I3)", 22.00, "=J3*K3"],
        ["James Okafor", "Cook", 8, 8, 0, 8, 8, 8, 0, "=SUM(C4:I4)", 17.00, "=J4*K4"],
        ["Priya Nair", "Server", 0, 6, 6, 6, 6, 8, 8, "=SUM(C5:I5)", 13.00, "=J5*K5"],
        ["Tyler Brooks", "Server", 6, 0, 6, 6, 6, 8, 8, "=SUM(C6:I6)", 13.00, "=J6*K6"],
        ["Aaliyah Grant", "Bartender", 0, 0, 5, 5, 8, 8, 8, "=SUM(C7:I7)", 15.00, "=J7*K7"],
        ["Leo Cheng", "Dishwasher", 6, 6, 6, 6, 0, 8, 8, "=SUM(C8:I8)", 12.50, "=J8*K8"],
        ["Sofia Reyes", "Host", 0, 5, 5, 5, 5, 8, 8, "=SUM(C9:I9)", 12.00, "=J9*K9"],
        ["Daniel Kim", "Cook", 8, 8, 8, 0, 0, 8, 8, "=SUM(C10:I10)", 17.00, "=J10*K10"],
        ["Fatima Al-Hassan", "Server", 5, 5, 0, 5, 8, 8, 0, "=SUM(C11:I11)", 13.00, "=J11*K11"],
        ["Chris Walters", "Manager", 0, 8, 8, 8, 8, 8, 0, "=SUM(C12:I12)", 22.00, "=J12*K12"],
    ]
    vals.append(("'Labor & Schedule'!A1", employees))
    vals.append(("'Labor & Schedule'!A22", [
        ["TOTALS", "", "", "", "", "", "", "", "", "=SUM(J3:J21)", "",
         "=SUM(L3:L21)"]
    ]))


def build_sheet5_checklist(sid: int, reqs: list, vals: list):
    """Sheet 5 – Opening & Closing Checklist"""
    print("  Building: Opening & Closing Checklist")

    DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

    # Opening section header (columns A-I, row 0)
    reqs.append(repeat_cell_req(sid, 0, 0, 1, 9,
        cell_fmt(bg=OPEN_HEADER, text=text_fmt(bold=True, size=14, fg=WHITE), h_align="CENTER"),
        "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"))

    # Closing section header (columns K-S, row 0) — offset by 1 blank col
    reqs.append(repeat_cell_req(sid, 0, 10, 1, 19,
        cell_fmt(bg=CLOSE_HEADER, text=text_fmt(bold=True, size=14, fg=WHITE), h_align="CENTER"),
        "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"))

    # Column sub-headers for both sections (row 1)
    reqs.append(repeat_cell_req(sid, 1, 0, 2, 9,
        cell_fmt(bg=OPEN_HEADER, text=text_fmt(bold=True, size=11, fg=WHITE), h_align="CENTER"),
        "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"))
    reqs.append(repeat_cell_req(sid, 1, 10, 2, 19,
        cell_fmt(bg=CLOSE_HEADER, text=text_fmt(bold=True, size=11, fg=WHITE), h_align="CENTER"),
        "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"))

    # Data rows
    for r in range(2, 10):
        bg = WHITE if r % 2 == 0 else LIGHT_GREY
        for start, end in [(0, 9), (10, 19)]:
            reqs.append(repeat_cell_req(sid, r, start, r + 1, end,
                cell_fmt(bg=bg, text=text_fmt(size=11), h_align="CENTER"),
                "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"))
        # Task name left-aligned
        reqs.append(repeat_cell_req(sid, r, 0, r + 1, 1,
            cell_fmt(text=text_fmt(bold=True, size=11), h_align="LEFT"),
            "userEnteredFormat(textFormat,horizontalAlignment)"))
        reqs.append(repeat_cell_req(sid, r, 10, r + 1, 11,
            cell_fmt(text=text_fmt(bold=True, size=11), h_align="LEFT"),
            "userEnteredFormat(textFormat,horizontalAlignment)"))

    reqs.append(freeze_req(sid, rows=2))
    reqs.append(tab_color_req(sid, OPEN_HEADER))
    reqs.append(auto_resize_req(sid, 0, 19))

    header_row = (
        ["OPENING TASKS"] + [""] * 8 +
        [""] +
        ["CLOSING TASKS"] + [""] * 8
    )
    sub_header = (
        ["Task", "Assigned To"] + DAYS +
        [""] +
        ["Task", "Assigned To"] + DAYS
    )
    opening_tasks = [
        "Unlock doors", "Check POS", "Staff check-in",
        "Prep stations", "Check inventory", "Turn on equipment",
    ]
    closing_tasks = [
        "Cash out register", "Clean kitchen", "Lock coolers",
        "Mop floors", "Set alarm", "Manager sign-off",
    ]
    rows = [header_row, sub_header]
    for i in range(6):
        o = opening_tasks[i] if i < len(opening_tasks) else ""
        c = closing_tasks[i] if i < len(closing_tasks) else ""
        row = [o, ""] + [""] * 7 + [""] + [c, ""] + [""] * 7
        rows.append(row)

    vals.append(("'Opening & Closing Checklist'!A1", rows))


def build_sheet6_temp_log(sid: int, reqs: list, vals: list):
    """Sheet 6 – Food Temp Log"""
    print("  Building: Food Temp Log")
    COLS = 8  # Date | Time | Item | Location | Temp (°F) | Safe? | Initials | Notes

    # Header
    reqs.append(repeat_cell_req(sid, 0, 0, 1, COLS,
        cell_fmt(bg=ROSE, text=text_fmt(bold=True, size=12, fg=DARK), h_align="CENTER"),
        "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"))

    # Alternating rows
    for r in range(1, 51):
        bg = WHITE if r % 2 == 1 else LIGHT_GREY
        reqs.append(repeat_cell_req(sid, r, 0, r + 1, COLS,
            cell_fmt(bg=bg, text=text_fmt(size=11)),
            "userEnteredFormat(backgroundColor,textFormat)"))

    # Safe? col conditional formatting
    reqs.append(cond_format_req(sid, 1, 5, 51, 6,
        '=F2="✓ SAFE"', GREEN_FILL, GREEN_TEXT))
    reqs.append(cond_format_req(sid, 1, 5, 51, 6,
        '=F2="⚠ CHECK"', RED_FILL, RED_TEXT))

    # Reference section header (row 52)
    reqs.append(repeat_cell_req(sid, 52, 0, 53, COLS,
        cell_fmt(bg=DARK, text=text_fmt(bold=True, size=12, fg=WHITE), h_align="CENTER"),
        "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"))

    # Reference data rows
    for r in range(53, 57):
        reqs.append(repeat_cell_req(sid, r, 0, r + 1, COLS,
            cell_fmt(bg=FAF0F0, text=text_fmt(size=11)),
            "userEnteredFormat(backgroundColor,textFormat)"))

    reqs.append(freeze_req(sid, rows=1))
    reqs.append(tab_color_req(sid, ROSE))
    reqs.append(auto_resize_req(sid, 0, COLS))

    sample = [
        ["Date", "Time", "Item", "Location", "Temp (°F)", "Safe?", "Initials", "Notes"],
        ["2024-04-01", "07:00", "Romaine Lettuce", "Fridge",     38,
         '=IF(AND(D2="Fridge",E2>=35,E2<=40),"✓ SAFE",IF(AND(D2="Freezer",E2<=0),"✓ SAFE","⚠ CHECK"))',
         "JS", ""],
        ["2024-04-01", "07:05", "Chicken Breast", "Fridge",     36,
         '=IF(AND(D3="Fridge",E3>=35,E3<=40),"✓ SAFE",IF(AND(D3="Freezer",E3<=0),"✓ SAFE","⚠ CHECK"))',
         "JS", ""],
        ["2024-04-01", "07:10", "Ice Cream",       "Freezer",   -5,
         '=IF(AND(D4="Fridge",E4>=35,E4<=40),"✓ SAFE",IF(AND(D4="Freezer",E4<=0),"✓ SAFE","⚠ CHECK"))',
         "MR", ""],
        ["2024-04-01", "07:15", "Beef Sirloin",    "Line",      41,
         '=IF(AND(D5="Fridge",E5>=35,E5<=40),"✓ SAFE",IF(AND(D5="Freezer",E5<=0),"✓ SAFE","⚠ CHECK"))',
         "MR", "Above safe fridge range"],
        ["2024-04-01", "11:00", "Salmon Fillet",   "Fridge",    37,
         '=IF(AND(D6="Fridge",E6>=35,E6<=40),"✓ SAFE",IF(AND(D6="Freezer",E6<=0),"✓ SAFE","⚠ CHECK"))',
         "AL", ""],
    ]
    vals.append(("'Food Temp Log'!A1", sample))

    # Reference section
    ref = [
        ["SAFE TEMPERATURE REFERENCE", "", "", "", "", "", "", ""],
        ["Location", "Min Temp (°F)", "Max Temp (°F)", "Formula Rule", "", "", "", ""],
        ["Fridge",   35,              40,              "35°F ≤ temp ≤ 40°F", "", "", "", ""],
        ["Freezer",  "< 0",           "N/A",           "temp < 0°F",         "", "", "", ""],
        ["Line",     "< 41",          "N/A",           "Below 41°F (danger zone starts at 41°F)", "", "", "", ""],
    ]
    vals.append(("'Food Temp Log'!A53", ref))


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

SHEET_DEFS = [
    {"title": "Dashboard",                  "builder": build_sheet1_dashboard},
    {"title": "Income & Expenses",          "builder": build_sheet2_income_expenses},
    {"title": "Food Inventory & Cost",      "builder": build_sheet3_inventory},
    {"title": "Labor & Schedule",           "builder": build_sheet4_labor},
    {"title": "Opening & Closing Checklist","builder": build_sheet5_checklist},
    {"title": "Food Temp Log",              "builder": build_sheet6_temp_log},
]


def main():
    print("=" * 60)
    print("Restaurant Monthly Budget & Operations Tracker")
    print("=" * 60)

    # ── Step 1: Create spreadsheet with all 6 sheets ──────────────
    print("\n[1/5] Creating spreadsheet...")
    create_payload = {
        "properties": {"title": "Restaurant Monthly Budget & Operations Tracker"},
        "sheets": [
            {"properties": {"title": s["title"], "index": i}}
            for i, s in enumerate(SHEET_DEFS)
        ],
    }
    result = run_gws(
        ["sheets", "spreadsheets", "create", "--json", json.dumps(create_payload)]
    )
    spreadsheet_id = result["spreadsheetId"]
    spreadsheet_url = result["spreadsheetUrl"]
    print(f"  ID:  {spreadsheet_id}")
    print(f"  URL: {spreadsheet_url}")

    # Build sheet_id map
    sheet_id_map = {}
    for sheet in result.get("sheets", []):
        props = sheet.get("properties", {})
        sheet_id_map[props["title"]] = props["sheetId"]

    print(f"  Sheets: {list(sheet_id_map.keys())}")

    # ── Step 2: Collect all formatting requests ───────────────────
    print("\n[2/5] Building formatting requests...")
    all_requests = []
    all_values = []

    for s in SHEET_DEFS:
        sid = sheet_id_map[s["title"]]
        s["builder"](sid, all_requests, all_values)

    # ── Step 3: Apply all formatting in one batchUpdate ───────────
    print(f"\n[3/5] Applying formatting ({len(all_requests)} requests)...")
    # Split into batches of 100 to avoid payload limits
    BATCH_SIZE = 100
    for i in range(0, len(all_requests), BATCH_SIZE):
        batch = all_requests[i : i + BATCH_SIZE]
        batch_update(spreadsheet_id, batch,
                     f"Batch {i//BATCH_SIZE + 1}/{(len(all_requests)-1)//BATCH_SIZE + 1}")

    # ── Step 4: Write all values ──────────────────────────────────
    print(f"\n[4/5] Writing data ({len(all_values)} ranges)...")
    for range_str, values in all_values:
        values_update(spreadsheet_id, range_str, values,
                      f"Range: {range_str[:40]}")

    # ── Step 5: Final column auto-resize ─────────────────────────
    print("\n[5/5] Auto-resizing columns...")
    resize_reqs = [auto_resize_req(sid, 0, 20) for sid in sheet_id_map.values()]
    batch_update(spreadsheet_id, resize_reqs, "Auto-resize all sheets")

    print("\n" + "=" * 60)
    print("SUCCESS!")
    print(f"Spreadsheet URL: {spreadsheet_url}")
    print("=" * 60)
    return spreadsheet_url


if __name__ == "__main__":
    main()
