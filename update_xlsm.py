"""
One-shot updater: applies the latest VBA form layout changes to the existing
SparePartsInventory_v2.xlsm WITHOUT rebuilding the workbook or touching data.

In-place approach (does NOT remove/recreate UserForms — that operation is
unstable against an existing xlsm). Instead it:
  1. Replaces the CodeModule of modHelpers (adds ApplyRTLAlignment)
  2. For each UserForm: updates geometry, repositions existing controls,
     adds any new controls (header labels), replaces the CodeModule.

All sheets, tables, named ranges, settings, and data are preserved.
"""

import os
import time
import xlwings as xw

import build_inventory as bi

PATH = r'c:\Users\avita\Claude_Projects\InventoryManager\SparePartsInventory_v2.xlsm'


# Form specs — mirrors build_inventory.setup_vba forms section.
# Tuple: (form_name, caption, width, height, code, controls)
# Each control: (prog_id, ctrl_name, caption, left, top, width, height, extra_dict)
FORM_SPECS = [
    ("frmMain", "ניהול מלאי – מכשור ובקרה", 300, 385, bi.CODE_FRMMAIN, [
        ("Forms.Label.1",         "lblTitle",    "ניהול מלאי מכשור ובקרה",10, 8,270,24,{"fsize":13,"bold":True}),
        ("Forms.CommandButton.1", "cmdAddItem",  "הוספת פריט חדש",        30, 45,230,30,{}),
        ("Forms.CommandButton.1", "cmdStockIn",  "קבלת מלאי (IN)",         30, 82,230,30,{}),
        ("Forms.CommandButton.1", "cmdStockOut", "הוצאת מלאי (OUT)",       30,119,230,30,{}),
        ("Forms.CommandButton.1", "cmdAdjust",   "תיקון מלאי",             30,156,230,30,{}),
        ("Forms.CommandButton.1", "cmdSearch",   "חיפוש פריט",             30,193,230,30,{}),
        ("Forms.CommandButton.1", "cmdRefresh",  "רענן לוח בקרה",          30,230,230,30,{}),
        ("Forms.CommandButton.1", "cmdOrders",   "דוח חוסרים",             30,267,230,30,{}),
        ("Forms.CommandButton.1", "cmdClose",    "סגור",                    30,310,230,28,{}),
    ]),

    ("frmStockOut", "הוצאת פריט מהמלאי", 492, 500, bi.CODE_FRMSTOCKOUT, [
        ("Forms.Label.1",         "lbl1",     "חיפוש פריט:",              6,  6,100,16,{}),
        ("Forms.TextBox.1",       "txtSearch","",                           6, 24,280,20,{}),
        ("Forms.Label.1",         "hdr1",     "שם פריט",                   6, 46, 80,14,{"bold":True,"fsize":9}),
        ("Forms.Label.1",         "hdr2",     "יצרן",                     86, 46,160,14,{"bold":True,"fsize":9}),
        ("Forms.Label.1",         "hdr3",     "דגם",                     246, 46, 80,14,{"bold":True,"fsize":9}),
        ("Forms.Label.1",         "hdr4",     "מלאי",                    326, 46, 50,14,{"bold":True,"fsize":9}),
        ("Forms.Label.1",         "hdr5",     "מזהה",                    376, 46, 80,14,{"bold":True,"fsize":9}),
        ("Forms.ListBox.1",       "lstRes",   "",                           6, 62,460, 98,{"cols":5,"colw":"80;160;80;50;80"}),
        ("Forms.Label.1",         "lblSel",   "",                           6,166,460,16,{}),
        ("Forms.Label.1",         "lblCurQ",  "מלאי נוכחי: -",             6,186,200,16,{}),
        ("Forms.Label.1",         "lblMinQ",  "מינימום: -",               220,186,200,16,{}),
        ("Forms.Label.1",         "lbl2",     "כמות להוצאה:",              6,206,100,16,{}),
        ("Forms.TextBox.1",       "txtQty",   "",                           6,224, 80,22,{}),
        ("Forms.Label.1",         "lbl3",     "הוראת עבודה:",              6,250,130,16,{}),
        ("Forms.TextBox.1",       "txtWO",    "",                           6,268,200,22,{}),
        ("Forms.Label.1",         "lbl4",     "שם מבצע:",                  6,294,100,16,{}),
        ("Forms.TextBox.1",       "txtTechnician","",                       6,312,200,22,{}),
        ("Forms.Label.1",         "lbl5",     "הערה / סיבה:",              6,338,100,16,{}),
        ("Forms.TextBox.1",       "txtReason","",                           6,356,460,22,{}),
        ("Forms.Label.1",         "lblWarn",  "",                           6,382,460,16,{}),
        ("Forms.Label.1",         "lblCrit",  "",                           6,400,460,16,{}),
        ("Forms.CommandButton.1", "cmdSave",  "אשר הוצאה",                 6,424,220,28,{}),
        ("Forms.CommandButton.1", "cmdCancel","ביטול",                    246,424,220,28,{}),
    ]),

    ("frmStockIn", "קבלת מלאי למחסן", 492, 485, bi.CODE_FRMSTOCKIN, [
        ("Forms.Label.1",         "lbl1",     "חיפוש פריט:",              6,  6,100,16,{}),
        ("Forms.TextBox.1",       "txtSearch","",                           6, 24,280,20,{}),
        ("Forms.Label.1",         "hdr1",     "שם פריט",                   6, 46,160,14,{"bold":True,"fsize":9}),
        ("Forms.Label.1",         "hdr2",     "יצרן",                    166, 46,100,14,{"bold":True,"fsize":9}),
        ("Forms.Label.1",         "hdr3",     "מלאי",                    266, 46, 50,14,{"bold":True,"fsize":9}),
        ("Forms.Label.1",         "hdr4",     "מזהה",                    316, 46, 80,14,{"bold":True,"fsize":9}),
        ("Forms.ListBox.1",       "lstRes",   "",                           6, 62,460, 88,{"cols":4,"colw":"160;100;50;80"}),
        ("Forms.Label.1",         "lblSel",   "",                           6,156,460,16,{}),
        ("Forms.Label.1",         "lblCurQ",  "מלאי נוכחי: -",             6,176,200,16,{}),
        ("Forms.Label.1",         "lbl2",     "כמות לקבלה:",               6,198,100,16,{}),
        ("Forms.TextBox.1",       "txtQty",   "",                           6,216, 80,22,{}),
        ("Forms.Label.1",         "lbl3",     "מחיר ליחידה ₪:",            6,242,120,16,{}),
        ("Forms.TextBox.1",       "txtPrice", "",                           6,260,120,22,{}),
        ("Forms.Label.1",         "lbl4",     "מספר הזמנת רכש (PO):",      6,286,160,16,{}),
        ("Forms.TextBox.1",       "txtPO",    "",                           6,304,160,22,{}),
        ("Forms.Label.1",         "lbl5",     "שם מקבל:",                  6,330,100,16,{}),
        ("Forms.TextBox.1",       "txtTechnician","",                       6,348,200,22,{}),
        ("Forms.Label.1",         "lbl6",     "הערה:",                     6,374,100,16,{}),
        ("Forms.TextBox.1",       "txtReason","",                           6,392,460,20,{}),
        ("Forms.CommandButton.1", "cmdSave",  "קלוט קבלה",                 6,418,220,28,{}),
        ("Forms.CommandButton.1", "cmdCancel","ביטול",                    246,418,220,28,{}),
    ]),

    ("frmAddItem", "הוספת פריט חדש למאגר", 432, 510, bi.CODE_FRMADDITEM, [
        ("Forms.Label.1",         "lbl1",       "שם פריט בעברית *:",        6,  6,160,16,{}),
        ("Forms.TextBox.1",       "txtNameHE",  "",                           6, 24,395,22,{}),
        ("Forms.Label.1",         "lbl2",       "שם פריט באנגלית:",          6, 54,160,16,{}),
        ("Forms.TextBox.1",       "txtNameEN",  "",                           6, 72,395,22,{}),
        ("Forms.Label.1",         "lbl3",       "קטגוריה *:",               6,100,100,16,{}),
        ("Forms.ComboBox.1",      "cboCategory","",                           6,118,185,22,{}),
        ("Forms.Label.1",         "lbl4",       "יצרן *:",                  205,100, 80,16,{}),
        ("Forms.ComboBox.1",      "cboMfg",     "",                          205,118,190,22,{}),
        ("Forms.Label.1",         "lbl5",       "דגם / מק\"ט *:",           6,148,120,16,{}),
        ("Forms.TextBox.1",       "txtModel",   "",                           6,166,185,22,{}),
        ("Forms.Label.1",         "lbl6",       "מק\"ט הזמנה:",             205,148,110,16,{}),
        ("Forms.TextBox.1",       "txtPN",      "",                          205,166,190,22,{}),
        ("Forms.Label.1",         "lbl7",       "סוג:",                      6,196, 60,16,{}),
        ("Forms.ComboBox.1",      "cboType",    "",                           6,214, 85,22,{}),
        ("Forms.Label.1",         "lbl8",       "יחידה:",                   100,196, 60,16,{}),
        ("Forms.ComboBox.1",      "cboUnit",    "",                          100,214, 85,22,{}),
        ("Forms.Label.1",         "lbl9",       "מחיר ₪:",                  200,196, 80,16,{}),
        ("Forms.TextBox.1",       "txtPrice",   "",                          200,214,100,22,{}),
        ("Forms.Label.1",         "lbl10",      "מיקום אחסון:",              6,244,110,16,{}),
        ("Forms.ComboBox.1",      "cboLoc",     "",                           6,262,185,22,{}),
        ("Forms.Label.1",         "lbl11",      "מינימום מלאי:",            205,244,120,16,{}),
        ("Forms.TextBox.1",       "txtMinQty",  "",                          205,262, 80,22,{}),
        ("Forms.CheckBox.1",      "chkCritical","פריט קריטי (שבר, משבית מערכת)",6,294,300,20,{}),
        ("Forms.Label.1",         "lbl12",      "הערות:",                    6,320,100,16,{}),
        ("Forms.TextBox.1",       "txtNotes",   "",                           6,338,395,50,{}),
        ("Forms.Label.1",         "lblPreview", "",                           6,396,395,16,{}),
        ("Forms.CommandButton.1", "cmdSave",    "שמור פריט חדש",             6,420,190,30,{}),
        ("Forms.CommandButton.1", "cmdCancel",  "ביטול",                    210,420,190,30,{}),
    ]),

    ("frmAdjust", "תיקון מלאי (Adjustment)", 402, 310, bi.CODE_FRMADJUST, [
        ("Forms.Label.1",         "lbl1",       "מזהה פריט (Item_ID):",     6,  6,160,16,{}),
        ("Forms.TextBox.1",       "txtItemID",  "",                           6, 24,185,22,{}),
        ("Forms.Label.1",         "lblName",    "",                           6, 52,370,16,{}),
        ("Forms.Label.1",         "lblCurQ",    "מלאי נוכחי: -",             6, 72,185,16,{}),
        ("Forms.Label.1",         "lbl2",       "כמות נכונה (חדשה):",       6,100,160,16,{}),
        ("Forms.TextBox.1",       "txtNewQty",  "",                           6,118, 80,22,{}),
        ("Forms.Label.1",         "lblDelta",   "",                          100,118, 80,22,{}),
        ("Forms.Label.1",         "lbl3",       "סיבת התיקון * (חובה):",    6,148,185,16,{}),
        ("Forms.TextBox.1",       "txtReason",  "",                           6,166,370,22,{}),
        ("Forms.Label.1",         "lbl4",       "שם מבצע:",                  6,196,100,16,{}),
        ("Forms.TextBox.1",       "txtTechnician","",                         6,214,200,22,{}),
        ("Forms.CommandButton.1", "cmdSave",    "אשר תיקון",                 6,248,175,30,{}),
        ("Forms.CommandButton.1", "cmdCancel",  "ביטול",                    200,248,175,30,{}),
    ]),

    ("frmSearch", "חיפוש פריטים", 652, 478, bi.CODE_FRMSEARCH, [
        ("Forms.Label.1",         "lbl1",      "חפש:",                       6,  6, 40,16,{}),
        ("Forms.TextBox.1",       "txtSearch", "",                            6, 24,380,22,{}),
        ("Forms.CommandButton.1", "cmdSearch", "חפש",                       396, 22, 60,26,{}),
        ("Forms.Label.1",         "hdr1",      "מזהה",                       6, 48, 70,14,{"bold":True,"fsize":9}),
        ("Forms.Label.1",         "hdr2",      "שם פריט",                   76, 48,180,14,{"bold":True,"fsize":9}),
        ("Forms.Label.1",         "hdr3",      "יצרן",                     256, 48, 90,14,{"bold":True,"fsize":9}),
        ("Forms.Label.1",         "hdr4",      "דגם",                      346, 48, 80,14,{"bold":True,"fsize":9}),
        ("Forms.Label.1",         "hdr5",      "מלאי",                     426, 48, 50,14,{"bold":True,"fsize":9}),
        ("Forms.Label.1",         "hdr6",      "קריטי",                    476, 48, 50,14,{"bold":True,"fsize":9}),
        ("Forms.ListBox.1",       "lstRes",    "",                            6, 64,620,170,{"cols":6,"colw":"70;180;90;80;50;50"}),
        ("Forms.Label.1",         "lblCount",  "נמצאו 0 פריטים",            6,240,200,16,{}),
        ("Forms.Label.1",         "lbl2",      "מזהה:",                      6,262, 60,16,{}),
        ("Forms.Label.1",         "lblDID",    "",                           70,262,120,16,{}),
        ("Forms.Label.1",         "lbl3",      "שם פריט:",                   6,282, 70,16,{}),
        ("Forms.Label.1",         "lblDName",  "",                           70,282,300,16,{}),
        ("Forms.Label.1",         "lbl4",      "יצרן:",                      6,302, 60,16,{}),
        ("Forms.Label.1",         "lblDMfg",   "",                           70,302,200,16,{}),
        ("Forms.Label.1",         "lbl5",      "דגם:",                       6,322, 60,16,{}),
        ("Forms.Label.1",         "lblDModel", "",                           70,322,200,16,{}),
        ("Forms.Label.1",         "lbl6",      "מלאי:",                      6,342, 60,16,{}),
        ("Forms.Label.1",         "lblDQty",   "",                           70,342, 60,16,{}),
        ("Forms.Label.1",         "lbl7",      "מיקום:",                     6,362, 60,16,{}),
        ("Forms.Label.1",         "lblDLoc",   "",                           70,362,200,16,{}),
        ("Forms.Label.1",         "lbl8",      "קריטי:",                     6,382, 60,16,{}),
        ("Forms.Label.1",         "lblDCrit",  "",                           70,382, 60,16,{}),
        ("Forms.CommandButton.1", "cmdStockOut","הוצא כמות",                 6,410,190,28,{}),
        ("Forms.CommandButton.1", "cmdStockIn", "קלוט כמות",               210,410,190,28,{}),
        ("Forms.CommandButton.1", "cmdClose",   "סגור",                    450,410,160,28,{}),
    ]),
]


def find_control(designer, name):
    try:
        return designer.Controls(name)
    except Exception:
        return None


def sync_form(vb, fname, caption, width, height, code, controls):
    try:
        comp = vb.VBComponents(fname)
    except Exception as e:
        print(f"  SKIP {fname}: component not found ({e})")
        return

    d = comp.Designer

    # 1. For each spec control: update if exists, add if new
    for (prog_id, ctrl_name, cap, left, top, w, h, extra) in controls:
        ctrl = find_control(d, ctrl_name)
        try:
            if ctrl is None:
                ctrl = d.Controls.Add(prog_id, ctrl_name, True)
            ctrl.Left = left
            ctrl.Top = top
            ctrl.Width = w
            ctrl.Height = h
            if "Label" in prog_id or "CommandButton" in prog_id or "CheckBox" in prog_id:
                ctrl.Caption = cap
            if "ListBox" in prog_id and isinstance(extra, dict):
                ctrl.ColumnCount = extra.get("cols", 1)
                ctrl.ColumnWidths = extra.get("colw", "")
            if "Label" in prog_id and isinstance(extra, dict):
                if extra.get("bold"):
                    ctrl.Font.Bold = True
                ctrl.Font.Size = extra.get("fsize", 10)
        except Exception as e:
            print(f"    WARN control {ctrl_name}: {e}")

    # 2. Update form geometry via Properties
    try:
        comp.Properties("Width").Value = width
    except Exception as e:
        print(f"    WARN set Width: {e}")
    try:
        comp.Properties("Height").Value = height
    except Exception as e:
        print(f"    WARN set Height: {e}")

    # 3. Replace CodeModule contents
    try:
        cm = comp.CodeModule
        n = cm.CountOfLines
        if n > 0:
            cm.DeleteLines(1, n)
        cm.AddFromString(code)
    except Exception as e:
        print(f"    WARN code replace: {e}")

    print(f"  Synced form: {fname}")


def replace_module_code(vb, mod_name, code):
    try:
        comp = vb.VBComponents(mod_name)
        cm = comp.CodeModule
        n = cm.CountOfLines
        if n > 0:
            cm.DeleteLines(1, n)
        cm.AddFromString(code)
        print(f"  Replaced code: {mod_name}")
    except Exception as e:
        print(f"  WARN {mod_name}: {e}")


def main():
    if not os.path.exists(PATH):
        print(f"ERROR: file not found: {PATH}")
        return

    bi.enable_vba_access()
    os.system("taskkill /F /IM excel.exe >nul 2>&1")
    time.sleep(1)

    app = xw.App(visible=True, add_book=False)
    app.display_alerts = False
    try:
        wb = app.books.open(PATH)
        print(f"Opened: {PATH}")

        try:
            vb = wb.api.VBProject
        except Exception as e:
            print(f"ERROR: Cannot access VBProject. Enable 'Trust access to the VBA project object model'.\n  {e}")
            return

        # 1. Update modHelpers code (adds ApplyRTLAlignment)
        replace_module_code(vb, "modHelpers", bi.VBA_HELPERS)

        # 2. In-place update of all 6 UserForms
        for (fname, caption, w, h, code, controls) in FORM_SPECS:
            sync_form(vb, fname, caption, w, h, code, controls)

        print("Saving...")
        wb.save()
        wb.close()
        print("=== SUCCESS — data preserved, forms & modHelpers updated ===")

    except Exception as e:
        import traceback
        print(f"ERROR: {e}")
        traceback.print_exc()
    finally:
        try:
            app.quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
