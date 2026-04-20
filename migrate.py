#!/usr/bin/env python3
"""
Direct Python migration: reads source xlsx, writes to SparePartsInventory_v2.xlsm
"""
import sys, os, time, winreg
from datetime import datetime

sys.stdout.reconfigure(line_buffering=True)

TARGET = r'c:\Users\avita\Claude_Projects\InventoryManager\SparePartsInventory_v2.xlsm'
SOURCE = r'c:\Users\avita\Claude_Projects\InventoryManager\מהדורה -1 ניהול מלאי חלקי חילוף מכשור.xlsx'

def enable_vba():
    try:
        k = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Office\16.0\Excel\Security", 0, winreg.KEY_SET_VALUE)
        winreg.SetValueEx(k, "AccessVBOM", 0, winreg.REG_DWORD, 1)
        winreg.CloseKey(k)
    except: pass

def safe(v):
    if v is None: return ""
    s = str(v).strip()
    return s if s != "None" else ""

def safe_num(v, default=0):
    try:
        f = float(str(v).replace(",",""))
        return f
    except: return default

def read_source_sheet6():
    import openpyxl
    wb = openpyxl.load_workbook(SOURCE, read_only=True, data_only=True)
    ws = wb.worksheets[5]
    items = []
    empty_streak = 0
    for r in ws.iter_rows(min_row=5, values_only=True):
        name   = safe(r[2])   # col C - name HE
        desc   = safe(r[4])   # col E - description/name EN
        itype  = safe(r[3])   # col D - INS/CON
        model  = safe(r[5])   # col F - model
        sup1   = safe(r[6])   # col G - supplier 1
        sup2   = safe(r[7])   # col H - supplier 2
        mfg    = safe(r[8])   # col I - manufacturer
        partno = safe(r[9])   # col J - part number
        price  = safe_num(r[10])  # col K - price
        qty    = int(safe_num(r[11]))  # col L - qty on hand
        minqty = int(safe_num(r[12])) # col M - min qty
        loc    = safe(r[13])  # col N - location
        svc    = safe(r[14])  # col O - service agreement
        notes  = safe(r[15])  # col P - notes

        if not name:
            empty_streak += 1
            if empty_streak >= 5:
                break
            continue
        empty_streak = 0
        if not name or (not mfg and not model):
            continue

        itype_norm = "CON" if itype.upper() in ("CON","MGCON","M","מתכלה") else "INS"
        items.append({
            "name": name, "desc": desc, "itype": itype_norm,
            "model": model, "sup1": sup1, "sup2": sup2, "mfg": mfg,
            "partno": partno, "price": price, "qty": qty, "minqty": minqty,
            "loc": loc, "svc": svc, "notes": notes
        })
    wb.close()
    return items

def migrate():
    enable_vba()
    os.system("taskkill /F /IM excel.exe >nul 2>&1")
    time.sleep(1)

    import xlwings as xw
    app = xw.App(visible=True, add_book=False)
    app.display_alerts = False
    try: app.api.Calculation = -4135  # manual
    except: pass

    print(f"Opening {os.path.basename(TARGET)}...")
    app.api.EnableEvents = False   # prevent Workbook_Open from firing
    wb = app.books.open(TARGET)
    time.sleep(2)                  # let Excel finish loading
    app.api.EnableEvents = True

    ws_items = wb.sheets["Items_Master"].api
    ws_inv   = wb.sheets["Inventory"].api
    ws_txn   = wb.sheets["Transactions"].api
    ws_set   = wb.sheets["Settings"].api

    tbl_items = ws_items.ListObjects("tbl_Items")
    tbl_inv   = ws_inv.ListObjects("tbl_Inventory")
    tbl_txn   = ws_txn.ListObjects("tbl_Transactions")

    # Read current sequence counter
    try:
        item_seq = int(wb.api.Names("cfg_NextItemSeq").RefersToRange.Value or 1)
        txn_seq  = int(wb.api.Names("cfg_NextTxnSeq").RefersToRange.Value or 1)
    except:
        item_seq = 1
        txn_seq  = 1

    year = datetime.now().year
    now  = datetime.now()

    print("Reading source data...")
    items = read_source_sheet6()
    print(f"  Found {len(items)} items in source")

    # Build set of existing Mfg+Model to avoid duplicates
    existing = set()
    if tbl_items.DataBodyRange is not None:
        n = tbl_items.ListColumns.Count
        mfg_col = tbl_items.ListColumns("Manufacturer").Index
        mdl_col = tbl_items.ListColumns("Model").Index
        for row in tbl_items.DataBodyRange.Rows:
            m = safe(row.Cells(1, mfg_col).Value)
            d = safe(row.Cells(1, mdl_col).Value)
            existing.add((m.upper(), d.upper()))

    print(f"  Existing items in target: {len(existing)}")

    added_items = 0
    txn_rows = []

    for item in items:
        key = (item["mfg"].upper(), item["model"].upper())
        if key in existing:
            continue

        item_id = f"ITM-{year}-{item_seq:04d}"
        item_seq += 1
        existing.add(key)

        # Add row to tbl_Items
        if tbl_items.DataBodyRange is None:
            new_row = tbl_items.ListRows.Add()
        else:
            new_row = tbl_items.ListRows.Add()

        def col(name):
            return tbl_items.ListColumns(name).Index

        r = new_row.Range
        r.Cells(1, col("Item_ID")).Value       = item_id
        r.Cells(1, col("Item_Name_HE")).Value  = item["name"]
        r.Cells(1, col("Item_Name_EN")).Value  = item["desc"]
        r.Cells(1, col("Category")).Value      = "Other"
        r.Cells(1, col("Manufacturer")).Value  = item["mfg"]
        r.Cells(1, col("Model")).Value         = item["model"]
        r.Cells(1, col("Part_Number")).Value   = item["partno"]
        r.Cells(1, col("Item_Type")).Value     = item["itype"]
        r.Cells(1, col("Unit")).Value          = "יח'"
        r.Cells(1, col("Unit_Price_ILS")).Value = item["price"] if item["price"] else None
        r.Cells(1, col("Storage_Location")).Value = item["loc"]
        r.Cells(1, col("Supplier1_ID")).Value  = item["sup1"]
        r.Cells(1, col("Supplier2_ID")).Value  = item["sup2"]
        r.Cells(1, col("Is_Critical")).Value   = "No"
        r.Cells(1, col("Status")).Value        = "Active"
        r.Cells(1, col("Notes")).Value         = item["notes"]
        r.Cells(1, col("Created_Date")).Value  = now
        r.Cells(1, col("Created_By")).Value    = os.environ.get("USERNAME","Migration")
        r.Cells(1, col("Modified_Date")).Value = now

        # Add row to tbl_Inventory
        inv_row = tbl_inv.ListRows.Add()
        def icol(name):
            return tbl_inv.ListColumns(name).Index
        ir = inv_row.Range
        ir.Cells(1, icol("Item_ID")).Value         = item_id
        ir.Cells(1, icol("Item_Name_HE")).Value    = item["name"]
        ir.Cells(1, icol("Category")).Value        = "Other"
        ir.Cells(1, icol("Manufacturer")).Value    = item["mfg"]
        ir.Cells(1, icol("Model")).Value           = item["model"]
        ir.Cells(1, icol("Qty_On_Hand")).Value     = item["qty"]
        ir.Cells(1, icol("Min_Qty")).Value         = item["minqty"]
        ir.Cells(1, icol("Is_Critical")).Value     = "No"
        ir.Cells(1, icol("Unit_Price_ILS")).Value  = item["price"] if item["price"] else None
        ir.Cells(1, icol("Last_Transaction_Date")).Value = now
        ir.Cells(1, icol("Last_Transaction_Type")).Value = "INITIAL"

        # Queue INITIAL transaction if qty > 0
        if item["qty"] > 0:
            txn_rows.append((item_id, item["name"], item["qty"], item["price"]))

        added_items += 1
        if added_items % 10 == 0:
            print(f"  Imported {added_items} items...")

    print(f"  Total new items added: {added_items}")

    # Write INITIAL transactions
    print(f"  Writing {len(txn_rows)} INITIAL transactions...")
    for item_id, name, qty, price in txn_rows:
        txn_id = f"TXN-{now.strftime('%Y%m%d')}-{txn_seq:04d}"
        txn_seq += 1
        txn_row = tbl_txn.ListRows.Add()
        def tcol(n): return tbl_txn.ListColumns(n).Index
        tr = txn_row.Range
        tr.Cells(1, tcol("Txn_ID")).Value         = txn_id
        tr.Cells(1, tcol("Txn_Date")).Value        = now
        tr.Cells(1, tcol("Txn_Type")).Value        = "INITIAL"
        tr.Cells(1, tcol("Item_ID")).Value         = item_id
        tr.Cells(1, tcol("Item_Name_HE")).Value    = name
        tr.Cells(1, tcol("Qty_Change")).Value      = qty
        tr.Cells(1, tcol("Qty_Before")).Value      = 0
        tr.Cells(1, tcol("Qty_After")).Value       = qty
        tr.Cells(1, tcol("Unit_Price_ILS")).Value  = price if price else None
        tr.Cells(1, tcol("Reason")).Value          = "פתיחת מערכת - ייבוא מהדורה 1"
        tr.Cells(1, tcol("Recorded_By")).Value     = os.environ.get("USERNAME","Migration")
        tr.Cells(1, tcol("Recorded_DateTime")).Value = now

    # Save sequence counters via named ranges in the workbook
    wb.api.Names("cfg_NextItemSeq").RefersToRange.Value = item_seq
    wb.api.Names("cfg_NextTxnSeq").RefersToRange.Value  = txn_seq

    try: app.api.Calculation = -4105  # xlCalculationAutomatic
    except: pass
    print("Saving workbook...")
    wb.save()
    print("=== MIGRATION COMPLETE ===")
    print(f"  New items:        {added_items}")
    print(f"  INITIAL txns:     {len(txn_rows)}")
    print(f"  Next Item seq:    {item_seq}")

if __name__ == "__main__":
    migrate()
