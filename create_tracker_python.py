#!/usr/bin/env python3
"""
Restaurant Monthly Budget & Operations Tracker — Google Sheets Creator
Pure Python using google-api-python-client (no gws CLI needed).

Usage:
    pip install google-api-python-client google-auth
    python3 create_tracker_python.py
"""

import json
import os
import sys

# ─── Install deps check ───────────────────────────────────────────────────────
try:
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request
except ImportError:
    print("Missing dependencies. Run:")
    print("  pip install google-api-python-client google-auth google-auth-httplib2")
    sys.exit(1)


# ─── Credentials (from env vars or credentials.json) ─────────────────────────
def load_credentials() -> Credentials:
    """
    Load credentials from environment variables or a JSON file.

    Set ONE of:
      - GOOGLE_REFRESH_TOKEN + GOOGLE_CLIENT_ID + GOOGLE_CLIENT_SECRET
      - GOOGLE_CREDENTIALS_FILE=/path/to/credentials.json
        (must have keys: refresh_token, client_id, client_secret)
      - GOOGLE_ACCESS_TOKEN  (short-lived; no auto-refresh)
    """
    creds_file = os.environ.get("GOOGLE_CREDENTIALS_FILE")
    if creds_file:
        with open(creds_file) as f:
            d = json.load(f)
        return Credentials(
            token=None,
            refresh_token=d["refresh_token"],
            token_uri="https://oauth2.googleapis.com/token",
            client_id=d["client_id"],
            client_secret=d["client_secret"],
            scopes=["https://www.googleapis.com/auth/spreadsheets"],
        )

    refresh_token = os.environ.get("GOOGLE_REFRESH_TOKEN")
    client_id     = os.environ.get("GOOGLE_CLIENT_ID")
    client_secret = os.environ.get("GOOGLE_CLIENT_SECRET")
    access_token  = os.environ.get("GOOGLE_ACCESS_TOKEN")

    if refresh_token and client_id and client_secret:
        return Credentials(
            token=None,
            refresh_token=refresh_token,
            token_uri="https://oauth2.googleapis.com/token",
            client_id=client_id,
            client_secret=client_secret,
            scopes=["https://www.googleapis.com/auth/spreadsheets"],
        )
    if access_token:
        return Credentials(token=access_token)

    print("ERROR: No credentials found.")
    print("Set GOOGLE_CREDENTIALS_FILE, or set all three of:")
    print("  GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET, GOOGLE_REFRESH_TOKEN")
    sys.exit(1)


# ─── Colour helpers ───────────────────────────────────────────────────────────
def c(h):
    h = h.lstrip("#")
    return {"red": int(h[0:2], 16)/255, "green": int(h[2:4], 16)/255, "blue": int(h[4:6], 16)/255}

DARK      = c("2C2C2C");  ROSE      = c("E8A0A0");  WHITE     = c("FFFFFF")
LGREY     = c("F2F2F2");  FAF0F0    = c("FAF0F0");  GRN_F     = c("C6EFCE")
GRN_T     = c("276221");  RED_F     = c("FFC7CE");  RED_T     = c("9C0006")
OPENHDR   = c("375623");  CLOSEHDR  = c("843C0C")


# ─── Request builders ─────────────────────────────────────────────────────────
def gr(sid, r1, c1, r2=None, c2=None):
    g = {"sheetId": sid, "startRowIndex": r1, "startColumnIndex": c1}
    if r2 is not None: g["endRowIndex"] = r2
    if c2 is not None: g["endColumnIndex"] = c2
    return g

def tfmt(bold=False, sz=11, fg=None, font="Arial"):
    f = {"fontFamily": font, "fontSize": sz, "bold": bold}
    if fg: f["foregroundColor"] = fg
    return f

def cfmt(bg=None, tf=None, ha=None, va=None, nf=None):
    f = {}
    if bg: f["backgroundColor"] = bg
    if tf: f["textFormat"] = tf
    if ha: f["horizontalAlignment"] = ha
    if va: f["verticalAlignment"] = va
    if nf: f["numberFormat"] = nf
    return f

def rep(sid, r1, c1, r2, c2, fmt, fields):
    return {"repeatCell": {"range": gr(sid,r1,c1,r2,c2),
                           "cell": {"userEnteredFormat": fmt}, "fields": fields}}

def merge(sid, r1, c1, r2, c2):
    return {"mergeCells": {"range": gr(sid,r1,c1,r2,c2), "mergeType": "MERGE_ALL"}}

def freeze(sid, rows=1, cols=0):
    return {"updateSheetProperties": {
        "properties": {"sheetId": sid,
                       "gridProperties": {"frozenRowCount": rows, "frozenColumnCount": cols}},
        "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount"}}

def tab_color(sid, color):
    return {"updateSheetProperties": {
        "properties": {"sheetId": sid, "tabColorStyle": {"rgbColor": color}},
        "fields": "tabColorStyle"}}

def auto_resize(sid, start=0, end=20):
    return {"autoResizeDimensions": {
        "dimensions": {"sheetId": sid, "dimension": "COLUMNS",
                       "startIndex": start, "endIndex": end}}}

def cond(sid, r1, c1, r2, c2, formula, bg, tc):
    # Conditional format rules only support: backgroundColor, textFormat with
    # bold/italic/strikethrough/foregroundColor (no fontSize or fontFamily).
    fmt = {}
    if bg: fmt["backgroundColor"] = bg
    if tc: fmt["textFormat"] = {"foregroundColor": tc, "bold": True}
    return {"addConditionalFormatRule": {
        "rule": {"ranges": [gr(sid,r1,c1,r2,c2)],
                 "booleanRule": {
                     "condition": {"type": "CUSTOM_FORMULA",
                                   "values": [{"userEnteredValue": formula}]},
                     "format": fmt}},
        "index": 0}}

CURRENCY = {"type": "CURRENCY", "pattern": "$#,##0.00"}
PERCENT  = {"type": "PERCENT",  "pattern": "0.00%"}
HDR_F    = "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)"
ALL_F    = "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
BG_TF    = "userEnteredFormat(backgroundColor,textFormat)"
NF       = "userEnteredFormat.numberFormat"
TF_F     = "userEnteredFormat.textFormat"
HA_F     = "userEnteredFormat(textFormat,horizontalAlignment)"


# ─── Sheet builders ───────────────────────────────────────────────────────────
def dashboard(sid, reqs, vals):
    reqs += [
        merge(sid, 0, 0, 1, 7),
        rep(sid,0,0,1,7, cfmt(bg=DARK, tf=tfmt(True,18,WHITE), ha="CENTER", va="MIDDLE"), HDR_F),
        rep(sid,1,0,2,7, cfmt(bg=ROSE, tf=tfmt(True,12,DARK),  ha="CENTER"), ALL_F),
        rep(sid,2,0,4,7, cfmt(bg=WHITE,tf=tfmt(sz=12),          ha="CENTER"), ALL_F),
        *[rep(sid,2,col,4,col+1, cfmt(nf=CURRENCY), NF) for col in [1,2,3]],
        *[rep(sid,2,col,4,col+1, cfmt(nf=PERCENT),  NF) for col in [4,5]],
        freeze(sid, rows=0), tab_color(sid, DARK), auto_resize(sid,0,7),
    ]
    vals += [("Dashboard!A1", [
        ["RESTAURANT MONTHLY OVERVIEW"],
        ["","Total Revenue","Total Expenses","Net Profit","Food Cost %","Labor Cost %",""],
        ["","='Income & Expenses'!D102","='Income & Expenses'!E102",
         "='Income & Expenses'!D102-'Income & Expenses'!E102",
         "='Food Inventory & Cost'!G52/Dashboard!C3",
         "='Labor & Schedule'!L22/Dashboard!C3",""],
    ])]


def income_expenses(sid, reqs, vals):
    COLS = 6
    reqs += [rep(sid,0,0,1,COLS, cfmt(bg=DARK,tf=tfmt(True,12,WHITE),ha="CENTER"), ALL_F)]
    for r in range(1,101):
        reqs.append(rep(sid,r,0,r+1,COLS, cfmt(bg=WHITE if r%2 else LGREY, tf=tfmt(sz=11)), BG_TF))
    for col in [3,4,5]:
        reqs.append(rep(sid,1,col,101,col+1, cfmt(nf=CURRENCY), NF))
    reqs += [
        rep(sid,101,0,102,COLS, cfmt(bg=DARK,tf=tfmt(True,12,WHITE),ha="CENTER"), ALL_F),
        *[rep(sid,101,col,102,col+1, cfmt(nf=CURRENCY), NF) for col in [3,4,5]],
        cond(sid,1,3,101,4, "=AND(D2<>\"\",D2>0)", GRN_F, GRN_T),
        cond(sid,1,4,101,5, "=AND(E2<>\"\",E2>0)", RED_F, RED_T),
        freeze(sid), tab_color(sid, GRN_F), auto_resize(sid,0,COLS),
    ]
    rows = [["Date","Category","Description","Income","Expense","Balance"],
            ["2024-04-01","Dine-In","Lunch service",1250.00,"","=IF(D3<>\"\",D3,0)-IF(E3<>\"\",E3,0)+F2"],
            ["2024-04-01","Food & Beverage","Produce order","",320.00,"=IF(D4<>\"\",D4,0)-IF(E4<>\"\",E4,0)+F3"],
            ["2024-04-01","Labor","Staff wages","",780.00,"=IF(D5<>\"\",D5,0)-IF(E5<>\"\",E5,0)+F4"],
            ["2024-04-02","Takeout","Online orders",430.00,"","=IF(D6<>\"\",D6,0)-IF(E6<>\"\",E6,0)+F5"],
            ["2024-04-02","Delivery","Third-party delivery",210.00,"","=IF(D7<>\"\",D7,0)-IF(E7<>\"\",E7,0)+F6"],
            ["2024-04-02","Utilities","Electric & gas","",195.00,"=IF(D8<>\"\",D8,0)-IF(E8<>\"\",E8,0)+F7"],
            ["2024-04-03","Bar Sales","Weekend bar",680.00,"","=IF(D9<>\"\",D9,0)-IF(E9<>\"\",E9,0)+F8"],
            ["2024-04-03","Rent","Monthly rent","",3200.00,"=IF(D10<>\"\",D10,0)-IF(E10<>\"\",E10,0)+F9"],
            ["2024-04-03","Marketing","Social media ads","",150.00,"=IF(D11<>\"\",D11,0)-IF(E11<>\"\",E11,0)+F10"],
            ["2024-04-04","Catering","Private event",1800.00,"","=IF(D12<>\"\",D12,0)-IF(E12<>\"\",E12,0)+F11"]]
    vals += [("'Income & Expenses'!A1", rows),
             ("'Income & Expenses'!A102", [["TOTALS","","","=SUM(D2:D101)","=SUM(E2:E101)","=D102-E102"]])]


def inventory(sid, reqs, vals):
    COLS = 8
    reqs += [rep(sid,0,0,1,COLS, cfmt(bg=DARK,tf=tfmt(True,12,WHITE),ha="CENTER"), ALL_F)]
    for r in range(1,51):
        reqs.append(rep(sid,r,0,r+1,COLS, cfmt(bg=WHITE if r%2 else FAF0F0, tf=tfmt(sz=11)), BG_TF))
    for col in [5,6]:
        reqs.append(rep(sid,1,col,51,col+1, cfmt(nf=CURRENCY), NF))
    reqs += [
        cond(sid,1,7,51,8, '=H2="REORDER"', RED_F, RED_T),
        cond(sid,1,7,51,8, '=H2="OK"',      GRN_F, GRN_T),
        freeze(sid), tab_color(sid, FAF0F0), auto_resize(sid,0,COLS),
    ]
    items = [["Item","Category","Unit","Par Level","Current Stock","Unit Cost","Total Value","Reorder?"],
             ["Romaine Lettuce","Produce","Case",3,2,28.00,"=E3*F3",'=IF(E3<D3,"REORDER","OK")'],
             ["Tomatoes","Produce","Case",2,3,22.00,"=E4*F4",'=IF(E4<D4,"REORDER","OK")'],
             ["Chicken Breast","Meat","lb",40,25,3.50,"=E5*F5",'=IF(E5<D5,"REORDER","OK")'],
             ["Beef Sirloin","Meat","lb",30,35,8.75,"=E6*F6",'=IF(E6<D6,"REORDER","OK")'],
             ["Salmon Fillet","Meat","lb",15,10,12.00,"=E7*F7",'=IF(E7<D7,"REORDER","OK")'],
             ["Cheddar Cheese","Dairy","lb",10,8,4.50,"=E8*F8",'=IF(E8<D8,"REORDER","OK")'],
             ["Heavy Cream","Dairy","qt",12,14,3.20,"=E9*F9",'=IF(E9<D9,"REORDER","OK")'],
             ["All-Purpose Flour","Dry Goods","lb",25,30,0.65,"=E10*F10",'=IF(E10<D10,"REORDER","OK")'],
             ["Pasta (Penne)","Dry Goods","lb",20,18,1.10,"=E11*F11",'=IF(E11<D11,"REORDER","OK")'],
             ["Olive Oil","Dry Goods","gal",5,4,18.00,"=E12*F12",'=IF(E12<D12,"REORDER","OK")'],
             ["White Wine","Beverages","bottle",12,15,9.00,"=E13*F13",'=IF(E13<D13,"REORDER","OK")'],
             ["Beer (Domestic)","Beverages","case",8,5,22.00,"=E14*F14",'=IF(E14<D14,"REORDER","OK")'],
             ["Sparkling Water","Beverages","case",6,7,14.00,"=E15*F15",'=IF(E15<D15,"REORDER","OK")'],
             ["Dish Soap","Cleaning Supplies","gal",4,3,8.50,"=E16*F16",'=IF(E16<D16,"REORDER","OK")'],
             ["Sanitizer Solution","Cleaning Supplies","gal",5,6,12.00,"=E17*F17",'=IF(E17<D17,"REORDER","OK")']]
    vals.append(("'Food Inventory & Cost'!A1", items))


def labor(sid, reqs, vals):
    COLS = 12
    reqs += [rep(sid,0,0,1,COLS, cfmt(bg=DARK,tf=tfmt(True,12,WHITE),ha="CENTER"), ALL_F)]
    for r in range(1,21):
        reqs.append(rep(sid,r,0,r+1,COLS, cfmt(bg=WHITE if r%2 else LGREY, tf=tfmt(sz=11)), BG_TF))
    reqs += [
        rep(sid,1,1,21,2,  cfmt(tf=tfmt(bold=True,sz=11)), TF_F),
        rep(sid,1,11,21,12, cfmt(nf=CURRENCY), NF),
        rep(sid,21,0,22,COLS, cfmt(bg=DARK,tf=tfmt(True,12,WHITE),ha="CENTER"), ALL_F),
        rep(sid,21,11,22,12, cfmt(nf=CURRENCY), NF),
        freeze(sid), tab_color(sid, c("4472C4")), auto_resize(sid,0,COLS),
    ]
    emps = [["Employee Name","Role","Mon","Tue","Wed","Thu","Fri","Sat","Sun","Total Hours","Hourly Rate","Weekly Cost"],
            ["Maria Santos","Manager",8,8,8,8,8,0,0,"=SUM(C3:I3)",22.00,"=J3*K3"],
            ["James Okafor","Cook",8,8,0,8,8,8,0,"=SUM(C4:I4)",17.00,"=J4*K4"],
            ["Priya Nair","Server",0,6,6,6,6,8,8,"=SUM(C5:I5)",13.00,"=J5*K5"],
            ["Tyler Brooks","Server",6,0,6,6,6,8,8,"=SUM(C6:I6)",13.00,"=J6*K6"],
            ["Aaliyah Grant","Bartender",0,0,5,5,8,8,8,"=SUM(C7:I7)",15.00,"=J7*K7"],
            ["Leo Cheng","Dishwasher",6,6,6,6,0,8,8,"=SUM(C8:I8)",12.50,"=J8*K8"],
            ["Sofia Reyes","Host",0,5,5,5,5,8,8,"=SUM(C9:I9)",12.00,"=J9*K9"],
            ["Daniel Kim","Cook",8,8,8,0,0,8,8,"=SUM(C10:I10)",17.00,"=J10*K10"],
            ["Fatima Al-Hassan","Server",5,5,0,5,8,8,0,"=SUM(C11:I11)",13.00,"=J11*K11"],
            ["Chris Walters","Manager",0,8,8,8,8,8,0,"=SUM(C12:I12)",22.00,"=J12*K12"]]
    vals += [("'Labor & Schedule'!A1", emps),
             ("'Labor & Schedule'!A22", [["TOTALS","","","","","","","","","=SUM(J3:J21)","","=SUM(L3:L21)"]])]


def checklist(sid, reqs, vals):
    DAYS = ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]
    reqs += [
        rep(sid,0,0,1,9,   cfmt(bg=OPENHDR, tf=tfmt(True,14,WHITE), ha="CENTER"), ALL_F),
        rep(sid,0,10,1,19,  cfmt(bg=CLOSEHDR,tf=tfmt(True,14,WHITE), ha="CENTER"), ALL_F),
        rep(sid,1,0,2,9,   cfmt(bg=OPENHDR, tf=tfmt(True,11,WHITE), ha="CENTER"), ALL_F),
        rep(sid,1,10,2,19,  cfmt(bg=CLOSEHDR,tf=tfmt(True,11,WHITE), ha="CENTER"), ALL_F),
    ]
    for r in range(2,10):
        bg = WHITE if r%2==0 else LGREY
        for s,e in [(0,9),(10,19)]:
            reqs.append(rep(sid,r,s,r+1,e, cfmt(bg=bg,tf=tfmt(sz=11),ha="CENTER"), ALL_F))
        reqs.append(rep(sid,r,0,r+1,1,   cfmt(tf=tfmt(bold=True,sz=11),ha="LEFT"), HA_F))
        reqs.append(rep(sid,r,10,r+1,11, cfmt(tf=tfmt(bold=True,sz=11),ha="LEFT"), HA_F))
    reqs += [freeze(sid,rows=2), tab_color(sid, OPENHDR), auto_resize(sid,0,19)]
    otasks = ["Unlock doors","Check POS","Staff check-in","Prep stations","Check inventory","Turn on equipment"]
    ctasks = ["Cash out register","Clean kitchen","Lock coolers","Mop floors","Set alarm","Manager sign-off"]
    rows = [
        ["OPENING TASKS"]+[""]*8+[""]+["CLOSING TASKS"]+[""]*8,
        ["Task","Assigned To"]+DAYS+[""]+["Task","Assigned To"]+DAYS,
    ]
    for i in range(6):
        rows.append([otasks[i],""] + [""]*7 + [""] + [ctasks[i],""] + [""]*7)
    vals.append(("'Opening & Closing Checklist'!A1", rows))


def temp_log(sid, reqs, vals):
    COLS = 8
    reqs += [rep(sid,0,0,1,COLS, cfmt(bg=ROSE,tf=tfmt(True,12,DARK),ha="CENTER"), ALL_F)]
    for r in range(1,51):
        reqs.append(rep(sid,r,0,r+1,COLS, cfmt(bg=WHITE if r%2 else LGREY, tf=tfmt(sz=11)), BG_TF))
    reqs += [
        cond(sid,1,5,51,6, '=F2="✓ SAFE"',  GRN_F, GRN_T),
        cond(sid,1,5,51,6, '=F2="⚠ CHECK"', RED_F, RED_T),
        rep(sid,52,0,53,COLS, cfmt(bg=DARK,tf=tfmt(True,12,WHITE),ha="CENTER"), ALL_F),
        *[rep(sid,r,0,r+1,COLS, cfmt(bg=FAF0F0,tf=tfmt(sz=11)), BG_TF) for r in range(53,57)],
        freeze(sid), tab_color(sid, ROSE), auto_resize(sid,0,COLS),
    ]
    F = '=IF(AND(D{r}="Fridge",E{r}>=35,E{r}<=40),"✓ SAFE",IF(AND(D{r}="Freezer",E{r}<=0),"✓ SAFE","⚠ CHECK"))'
    sample = [
        ["Date","Time","Item","Location","Temp (°F)","Safe?","Initials","Notes"],
        ["2024-04-01","07:00","Romaine Lettuce","Fridge",38, F.format(r=2),"JS",""],
        ["2024-04-01","07:05","Chicken Breast", "Fridge",36, F.format(r=3),"JS",""],
        ["2024-04-01","07:10","Ice Cream",      "Freezer",-5,F.format(r=4),"MR",""],
        ["2024-04-01","07:15","Beef Sirloin",   "Line",  41, F.format(r=5),"MR","Above safe fridge range"],
        ["2024-04-01","11:00","Salmon Fillet",  "Fridge",37, F.format(r=6),"AL",""],
    ]
    ref = [
        ["SAFE TEMPERATURE REFERENCE","","","","","","",""],
        ["Location","Min Temp (°F)","Max Temp (°F)","Formula Rule","","","",""],
        ["Fridge",35,40,"35°F ≤ temp ≤ 40°F","","","",""],
        ["Freezer","< 0","N/A","temp < 0°F","","","",""],
        ["Line","< 41","N/A","Below 41°F","","","",""],
    ]
    vals += [("'Food Temp Log'!A1", sample), ("'Food Temp Log'!A53", ref)]


# ─── Main ─────────────────────────────────────────────────────────────────────
SHEET_DEFS = [
    {"title": "Dashboard",                   "fn": dashboard},
    {"title": "Income & Expenses",           "fn": income_expenses},
    {"title": "Food Inventory & Cost",       "fn": inventory},
    {"title": "Labor & Schedule",            "fn": labor},
    {"title": "Opening & Closing Checklist", "fn": checklist},
    {"title": "Food Temp Log",               "fn": temp_log},
]

def main():
    print("=" * 60)
    print("Restaurant Monthly Budget & Operations Tracker")
    print("=" * 60)

    print("\nAuthenticating...")
    creds = load_credentials()
    if not creds.token:
        creds.refresh(Request())
    service = build("sheets", "v4", credentials=creds)
    ss = service.spreadsheets()

    # Create spreadsheet with all sheets
    print("[1/4] Creating spreadsheet...")
    result = ss.create(body={
        "properties": {"title": "Restaurant Monthly Budget & Operations Tracker"},
        "sheets": [{"properties": {"title": s["title"], "index": i}}
                   for i, s in enumerate(SHEET_DEFS)],
    }).execute()
    sid_map = {sh["properties"]["title"]: sh["properties"]["sheetId"]
               for sh in result["sheets"]}
    url = result["spreadsheetUrl"]
    print(f"  URL: {url}")

    # Build all requests and value ranges
    print("[2/4] Building formatting requests...")
    all_reqs, all_vals = [], []
    for s in SHEET_DEFS:
        s["fn"](sid_map[s["title"]], all_reqs, all_vals)

    # Apply formatting in batches
    print(f"[3/4] Applying formatting ({len(all_reqs)} requests)...")
    BS = 100
    for i in range(0, len(all_reqs), BS):
        ss.batchUpdate(
            spreadsheetId=result["spreadsheetId"],
            body={"requests": all_reqs[i:i+BS]},
        ).execute()
        print(f"  Batch {i//BS+1}/{(len(all_reqs)-1)//BS+1}")

    # Write data
    print(f"[4/4] Writing data ({len(all_vals)} ranges)...")
    for rng, values in all_vals:
        ss.values().update(
            spreadsheetId=result["spreadsheetId"],
            range=rng,
            valueInputOption="USER_ENTERED",
            body={"values": values},
        ).execute()
        print(f"  {rng[:50]}")

    print("\n" + "=" * 60)
    print("SUCCESS!")
    print(f"Spreadsheet URL: {url}")
    print("=" * 60)
    return url


if __name__ == "__main__":
    main()
