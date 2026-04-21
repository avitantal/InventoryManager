#!/usr/bin/env python3
"""
SparePartsInventory_v2.xlsm Builder  –  uses xlwings for robust Excel automation
Industrial I&C Spare Parts Inventory Management System
"""
import sys, os, time, winreg
import xlwings as xw

OUTPUT_DIR  = r"c:\Users\avita\Claude_Projects\InventoryManager"
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "SparePartsInventory_v2.xlsm")

def rgb(r, g, b): return r + g*256 + b*65536

SHEET_DEFS = [
    ("Dashboard",           rgb(68,114,196)),
    ("Items_Master",        rgb(112,173,71)),
    ("Inventory",           rgb(112,173,71)),
    ("Min_Stock",           rgb(112,173,71)),
    ("Transactions",        rgb(128,128,128)),
    ("Suppliers",           rgb(75,172,198)),
    ("Purchase_Followup",   rgb(75,172,198)),
    ("Assets_Link",         rgb(112,48,160)),
    ("Lists",               rgb(255,255,255)),
    ("Settings",            rgb(255,255,255)),
    ("Archive",             rgb(255,255,255)),
]

CATEGORIES = [
    "Transmitter","Sensor / Switch","Analyzer","Positioner / Actuator",
    "Control Valve","PLC / Controller","I/O Card / Module","HMI / Panel",
    "Power Supply","Comm Module","Relay / Solenoid","Instrument Accessory",
    "Electronic Control Part","Cable / Connector","Other"
]
MANUFACTURERS = [
    "ABB","ABB ISRAEL","E+H (Endress+Hauser)","E+H ISRAEL",
    "FESTO","FESTO ISRAEL","SIEMENS","LABOM","KAMA","GEA",
    "NEGELE","THORNTON","METTLER TOLEDO","MODCON","CONTROTEC",
    "INSTARMETRIX","MADID","CONTEL","UNITED","BECK","STONEWALL",
    "ROTEX","WIKA","VEGA","EMERSON","YOKOGAWA","HONEYWELL",
    "PEPPERL+FUCHS","TURCK","BALLUFF","PHOENIX CONTACT","MURR",
    "SICK","IFM","DANFOSS","BURKERT","VALTEK","METSO","SAMSON",
    "FISHER","NELES","BIFFI","FLOWSERVE","SPIRAX SARCO","Other"
]
LOCATIONS   = ["מחסן בקרה","מחסן בקרה - כיולים","גג / ROOF 2","מחסן אחזקה","מחסן כללי","שטח"]
TXN_TYPES   = ["IN","OUT","ADJUST","INITIAL"]
UNITS       = ["יח'","זוג","סט","מ\"מ","מטר","ליטר","ק\"ג"]
ITEM_STATUS = ["Active","Inactive","Obsolete"]
PO_STATUS   = ["Pending","Partial","Received","Cancelled"]
YESNO       = ["Yes","No"]
ITEM_TYPES  = ["INS","CON"]
REASONS     = ["תקלה בשטח","החלפה שוטפת","תחזוקה מונעת","קליטת רכש",
               "תיקון/החזרה","גריעה מהמלאי","תיקון מלאי","העברה בין מיקומים","בדיקה","אחר"]

H_ITEMS = ["Item_ID","Item_Name_HE","Item_Name_EN","Category","Sub_Category",
           "Manufacturer","Model","Part_Number","Item_Type","Unit",
           "Unit_Price_ILS","Lead_Time_Months","Storage_Location",
           "Supplier1_ID","Supplier2_ID","Datasheet_URL",
           "Is_Critical","Failure_Probability","Shutdown_Cost_Day",
           "Service_Agreement","Status","Notes",
           "Created_Date","Created_By","Modified_Date"]

H_INV  = ["Item_ID","Item_Name_HE","Category","Manufacturer","Model",
          "Qty_On_Hand","Min_Qty","Shortage_Flag","Is_Critical",
          "Alert_Level","Unit_Price_ILS","Stock_Value_ILS",
          "Last_Transaction_Date","Last_Transaction_Type"]

H_MIN  = ["Item_ID","Item_Name_HE","Min_Qty","Reorder_Qty",
          "Is_Critical","Shutdown_Cost_Day","Priority_Score",
          "Review_Date","Notes"]

H_TXN  = ["Txn_ID","Txn_Date","Txn_Type","Item_ID","Item_Name_HE",
          "Qty_Change","Qty_Before","Qty_After","Unit_Price_ILS",
          "Reason","PO_Number","Work_Order",
          "Technician","Recorded_By","Recorded_DateTime","Related_Asset_Tag"]

H_SUP  = ["Supplier_ID","Supplier_Name_HE","Supplier_Name_EN",
          "Contact_Person","Phone","Email",
          "Manufacturer_Represented","Lead_Time_Typical_Months","Notes","Status"]

H_PO   = ["PO_ID","PO_Date","Item_ID","Item_Name_HE","Supplier_ID",
          "Qty_Ordered","Qty_Received","Unit_Price_ILS","Total_Price_ILS",
          "Expected_Delivery","Actual_Delivery","PO_Status",
          "Priority","Requested_By","Notes"]

H_ASS  = ["Asset_Tag","Asset_Name_HE","System","Manufacturer","Model",
          "Is_Critical","Item_ID_1","Item_ID_2","Item_ID_3","Item_ID_4","Notes"]

H_ARC  = H_TXN + ["Archive_Date","Archive_Reason"]

# ════════════════════════════════════════════════════════════════════════════
#  VBA CODE STRINGS
# ════════════════════════════════════════════════════════════════════════════

VBA_HELPERS = """\
Option Explicit

Public Function SafeStr(v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then SafeStr = "" Else SafeStr = Trim(CStr(v))
End Function
Public Function SafeLng(v As Variant) As Long
    If IsNull(v) Or IsEmpty(v) Or Not IsNumeric(v) Then SafeLng = 0 Else SafeLng = CLng(v)
End Function
Public Function SafeCur(v As Variant) As Currency
    If IsNull(v) Or IsEmpty(v) Or Not IsNumeric(v) Then SafeCur = 0 Else SafeCur = CCur(v)
End Function
Public Function TblRow(lo As ListObject, colName As String, findVal As String) As Long
    On Error GoTo EH
    If lo.DataBodyRange Is Nothing Then GoTo EH
    Dim r As Variant
    r = Application.Match(findVal, lo.ListColumns(colName).DataBodyRange, 0)
    If IsError(r) Then GoTo EH
    TblRow = CLng(r)
    Exit Function
EH: TblRow = 0
End Function
Public Sub ShowErr(msg As String): MsgBox msg, vbCritical, "שגיאה": End Sub
Public Sub ShowOK(msg As String):  MsgBox msg, vbInformation, "הצלחה": End Sub

Public Sub ApplyRTLAlignment(frm As Object)
    Dim ctl As Object
    For Each ctl In frm.Controls
        On Error Resume Next
        Select Case TypeName(ctl)
            Case "Label", "TextBox", "ComboBox"
                ctl.TextAlign = 3
        End Select
        On Error GoTo 0
    Next ctl
End Sub
"""

VBA_ITEMS = """\
Option Explicit

Public Function NextItemID() As String
    Dim seq As Long
    seq = SafeLng(ThisWorkbook.Names("cfg_NextItemSeq").RefersToRange.Value)
    If seq < 1 Then seq = 1
    NextItemID = "ITM-" & Year(Now) & "-" & Format(seq, "0000")
    ThisWorkbook.Names("cfg_NextItemSeq").RefersToRange.Value = seq + 1
End Function

Public Function ItemByID(itemID As String) As Boolean
    Dim lo As ListObject
    Set lo = ThisWorkbook.Sheets("Items_Master").ListObjects("tbl_Items")
    If lo.DataBodyRange Is Nothing Then ItemByID = False: Exit Function
    ItemByID = Not IsError(Application.Match(itemID, lo.ListColumns("Item_ID").DataBodyRange, 0))
End Function

Public Function ItemByMfgModel(mfg As String, model As String) As String
    Dim lo As ListObject
    Set lo = ThisWorkbook.Sheets("Items_Master").ListObjects("tbl_Items")
    If lo.DataBodyRange Is Nothing Then ItemByMfgModel = "": Exit Function
    Dim i As Long
    For i = 1 To lo.ListRows.Count
        If LCase(SafeStr(lo.DataBodyRange(i, lo.ListColumns("Manufacturer").Index).Value)) = LCase(mfg) And _
           LCase(SafeStr(lo.DataBodyRange(i, lo.ListColumns("Model").Index).Value)) = LCase(model) Then
            ItemByMfgModel = SafeStr(lo.DataBodyRange(i, lo.ListColumns("Item_ID").Index).Value)
            Exit Function
        End If
    Next i
    ItemByMfgModel = ""
End Function

Public Function GetItemFld(itemID As String, fld As String) As Variant
    Dim lo As ListObject
    Set lo = ThisWorkbook.Sheets("Items_Master").ListObjects("tbl_Items")
    If lo.DataBodyRange Is Nothing Then GetItemFld = "": Exit Function
    Dim r As Long: r = TblRow(lo, "Item_ID", itemID)
    If r = 0 Then GetItemFld = "": Exit Function
    GetItemFld = lo.DataBodyRange(r, lo.ListColumns(fld).Index).Value
End Function

Public Function AddItem(nameHE As String, nameEN As String, cat As String, subCat As String, _
    mfg As String, model As String, partNum As String, itype As String, _
    unit As String, price As Currency, leadTime As Double, loc As String, _
    sup1 As String, sup2 As String, datasheet As String, _
    isCrit As String, failProb As Double, shutCost As Currency, _
    svcAgmt As String, status As String, notes As String, _
    minQty As Long, reorderQty As Long) As String

    If Len(Trim(nameHE)) = 0 Then ShowErr "שם פריט בעברית הוא חובה": AddItem = "": Exit Function
    If Len(Trim(mfg)) = 0    Then ShowErr "יצרן הוא חובה":            AddItem = "": Exit Function
    If Len(Trim(model)) = 0  Then ShowErr "דגם הוא חובה":             AddItem = "": Exit Function
    If Len(Trim(cat)) = 0    Then ShowErr "קטגוריה היא חובה":         AddItem = "": Exit Function
    If ItemByMfgModel(mfg, model) <> "" Then
        ShowErr "פריט דומה כבר קיים (יצרן+דגם): " & ItemByMfgModel(mfg, model)
        AddItem = "": Exit Function
    End If
    Dim newID As String: newID = NextItemID()
    Application.ScreenUpdating = False
    Dim lo As ListObject
    Set lo = ThisWorkbook.Sheets("Items_Master").ListObjects("tbl_Items")
    Dim nr As ListRow: Set nr = lo.ListRows.Add
    Dim r As Range:    Set r  = nr.Range
    With lo
        r.Cells(1,.ListColumns("Item_ID").Index).Value          = newID
        r.Cells(1,.ListColumns("Item_Name_HE").Index).Value     = nameHE
        r.Cells(1,.ListColumns("Item_Name_EN").Index).Value     = nameEN
        r.Cells(1,.ListColumns("Category").Index).Value         = cat
        r.Cells(1,.ListColumns("Sub_Category").Index).Value     = subCat
        r.Cells(1,.ListColumns("Manufacturer").Index).Value     = mfg
        r.Cells(1,.ListColumns("Model").Index).Value            = model
        r.Cells(1,.ListColumns("Part_Number").Index).Value      = partNum
        r.Cells(1,.ListColumns("Item_Type").Index).Value        = itype
        r.Cells(1,.ListColumns("Unit").Index).Value             = unit
        r.Cells(1,.ListColumns("Unit_Price_ILS").Index).Value   = price
        r.Cells(1,.ListColumns("Lead_Time_Months").Index).Value = leadTime
        r.Cells(1,.ListColumns("Storage_Location").Index).Value = loc
        r.Cells(1,.ListColumns("Supplier1_ID").Index).Value     = sup1
        r.Cells(1,.ListColumns("Supplier2_ID").Index).Value     = sup2
        r.Cells(1,.ListColumns("Datasheet_URL").Index).Value    = datasheet
        r.Cells(1,.ListColumns("Is_Critical").Index).Value      = isCrit
        If failProb  > 0 Then r.Cells(1,.ListColumns("Failure_Probability").Index).Value = failProb
        If shutCost  > 0 Then r.Cells(1,.ListColumns("Shutdown_Cost_Day").Index).Value   = shutCost
        r.Cells(1,.ListColumns("Service_Agreement").Index).Value = svcAgmt
        r.Cells(1,.ListColumns("Status").Index).Value           = status
        r.Cells(1,.ListColumns("Notes").Index).Value            = notes
        r.Cells(1,.ListColumns("Created_Date").Index).Value     = Now()
        r.Cells(1,.ListColumns("Created_By").Index).Value       = Environ("USERNAME")
    End With
    Dim loI As ListObject
    Set loI = ThisWorkbook.Sheets("Inventory").ListObjects("tbl_Inventory")
    Dim nrI As ListRow: Set nrI = loI.ListRows.Add
    nrI.Range.Cells(1,loI.ListColumns("Item_ID").Index).Value     = newID
    nrI.Range.Cells(1,loI.ListColumns("Qty_On_Hand").Index).Value = 0
    Dim loM As ListObject
    Set loM = ThisWorkbook.Sheets("Min_Stock").ListObjects("tbl_MinStock")
    Dim nrM As ListRow: Set nrM = loM.ListRows.Add
    nrM.Range.Cells(1,loM.ListColumns("Item_ID").Index).Value      = newID
    nrM.Range.Cells(1,loM.ListColumns("Min_Qty").Index).Value      = minQty
    nrM.Range.Cells(1,loM.ListColumns("Reorder_Qty").Index).Value  = reorderQty
    Application.ScreenUpdating = True
    AddItem = newID
End Function
"""

VBA_TXN = """\
Option Explicit

Public Function NextTxnID() As String
    Dim seq As Long
    seq = SafeLng(ThisWorkbook.Names("cfg_NextTxnSeq").RefersToRange.Value)
    If seq < 1 Then seq = 1
    NextTxnID = "TXN-" & Format(Now,"YYYYMMDD") & "-" & Format(seq,"0000")
    ThisWorkbook.Names("cfg_NextTxnSeq").RefersToRange.Value = seq + 1
End Function

Public Sub LogTxn(txType As String, itemID As String, _
    qDelta As Long, qBefore As Long, qAfter As Long, _
    price As Currency, reason As String, po As String, _
    wo As String, tech As String, assetTag As String)
    Dim lo As ListObject
    Set lo = ThisWorkbook.Sheets("Transactions").ListObjects("tbl_Transactions")
    Dim nr As ListRow: Set nr = lo.ListRows.Add
    Dim r As Range:    Set r  = nr.Range
    With lo
        r.Cells(1,.ListColumns("Txn_ID").Index).Value            = NextTxnID()
        r.Cells(1,.ListColumns("Txn_Date").Index).Value          = Now()
        r.Cells(1,.ListColumns("Txn_Type").Index).Value          = txType
        r.Cells(1,.ListColumns("Item_ID").Index).Value           = itemID
        r.Cells(1,.ListColumns("Item_Name_HE").Index).Value      = SafeStr(GetItemFld(itemID,"Item_Name_HE"))
        r.Cells(1,.ListColumns("Qty_Change").Index).Value        = qDelta
        r.Cells(1,.ListColumns("Qty_Before").Index).Value        = qBefore
        r.Cells(1,.ListColumns("Qty_After").Index).Value         = qAfter
        r.Cells(1,.ListColumns("Unit_Price_ILS").Index).Value    = price
        r.Cells(1,.ListColumns("Reason").Index).Value            = reason
        r.Cells(1,.ListColumns("PO_Number").Index).Value         = po
        r.Cells(1,.ListColumns("Work_Order").Index).Value        = wo
        r.Cells(1,.ListColumns("Technician").Index).Value        = tech
        r.Cells(1,.ListColumns("Recorded_By").Index).Value       = Environ("USERNAME")
        r.Cells(1,.ListColumns("Recorded_DateTime").Index).Value = Now()
        r.Cells(1,.ListColumns("Related_Asset_Tag").Index).Value = assetTag
    End With
End Sub
"""

VBA_INV = """\
Option Explicit

Public Function GetQty(itemID As String) As Long
    Dim lo As ListObject
    Set lo = ThisWorkbook.Sheets("Inventory").ListObjects("tbl_Inventory")
    If lo.DataBodyRange Is Nothing Then GetQty = 0: Exit Function
    Dim r As Long: r = TblRow(lo,"Item_ID",itemID)
    If r = 0 Then GetQty = 0: Exit Function
    GetQty = SafeLng(lo.DataBodyRange(r,lo.ListColumns("Qty_On_Hand").Index).Value)
End Function

Private Sub SetQty(itemID As String, newQ As Long, txType As String)
    Dim lo As ListObject
    Set lo = ThisWorkbook.Sheets("Inventory").ListObjects("tbl_Inventory")
    Dim r As Long: r = TblRow(lo,"Item_ID",itemID)
    If r = 0 Then Exit Sub
    lo.DataBodyRange(r,lo.ListColumns("Qty_On_Hand").Index).Value           = newQ
    lo.DataBodyRange(r,lo.ListColumns("Last_Transaction_Date").Index).Value = Now()
    lo.DataBodyRange(r,lo.ListColumns("Last_Transaction_Type").Index).Value = txType
End Sub

Public Sub StockIn(itemID As String, qty As Long, price As Currency, _
                   po As String, reason As String, tech As String)
    If Not ItemByID(itemID) Then ShowErr "פריט לא נמצא: " & itemID: Exit Sub
    If qty <= 0 Then ShowErr "כמות חייבת להיות > 0": Exit Sub
    Dim cur As Long: cur = GetQty(itemID)
    Application.ScreenUpdating = False
    SetQty itemID, cur+qty, "IN"
    LogTxn "IN",itemID,qty,cur,cur+qty,price,reason,po,"",tech,""
    If po <> "" Then UpdatePO po,qty
    modDashboard.RefreshAll
    Application.ScreenUpdating = True
End Sub

Public Sub StockOut(itemID As String, qty As Long, reason As String, _
                    wo As String, tech As String, assetTag As String)
    If Not ItemByID(itemID) Then ShowErr "פריט לא נמצא: " & itemID: Exit Sub
    If qty <= 0 Then ShowErr "כמות חייבת להיות > 0": Exit Sub
    Dim cur As Long: cur = GetQty(itemID)
    If qty > cur Then ShowErr "כמות (" & qty & ") גדולה מהמלאי (" & cur & ")": Exit Sub
    Dim newQ As Long: newQ = cur - qty
    Dim price As Currency: price = SafeCur(GetItemFld(itemID,"Unit_Price_ILS"))
    Application.ScreenUpdating = False
    SetQty itemID,newQ,"OUT"
    LogTxn "OUT",itemID,-qty,cur,newQ,price,reason,"",wo,tech,assetTag
    modDashboard.RefreshAll
    Application.ScreenUpdating = True
    Dim loM As ListObject: Set loM = ThisWorkbook.Sheets("Min_Stock").ListObjects("tbl_MinStock")
    Dim rm As Long: rm = TblRow(loM,"Item_ID",itemID)
    Dim minQ As Long: If rm > 0 Then minQ = SafeLng(loM.DataBodyRange(rm,loM.ListColumns("Min_Qty").Index).Value)
    Dim nm As String: nm = SafeStr(GetItemFld(itemID,"Item_Name_HE"))
    If newQ = 0 And SafeStr(GetItemFld(itemID,"Is_Critical")) = "Yes" Then
        MsgBox "אזהרה! פריט קריטי (שבר, משבית מערכת) הגיע לאפס:" & vbCrLf & nm & vbCrLf & "יש לפתוח הזמנה דחופה!",vbCritical,"חוסר קריטי"
    ElseIf newQ < minQ Then
        MsgBox "שים לב: " & nm & " ירד מתחת למינימום (" & newQ & " < " & minQ & ")",vbExclamation,"אזהרת מלאי"
    End If
End Sub

Public Sub StockAdjust(itemID As String, newQ As Long, reason As String, tech As String)
    If Len(Trim(reason)) = 0 Then ShowErr "סיבה היא חובה לתיקון מלאי": Exit Sub
    If Not ItemByID(itemID) Then ShowErr "פריט לא נמצא: " & itemID: Exit Sub
    If newQ < 0 Then ShowErr "כמות לא יכולה להיות שלילית": Exit Sub
    Dim cur As Long: cur = GetQty(itemID)
    If newQ = cur Then ShowErr "אין שינוי בכמות": Exit Sub
    Dim price As Currency: price = SafeCur(GetItemFld(itemID,"Unit_Price_ILS"))
    Application.ScreenUpdating = False
    SetQty itemID,newQ,"ADJUST"
    LogTxn "ADJUST",itemID,newQ-cur,cur,newQ,price,reason,"","",tech,""
    modDashboard.RefreshAll
    Application.ScreenUpdating = True
End Sub

Private Sub UpdatePO(poID As String, qtyRec As Long)
    Dim lo As ListObject
    Set lo = ThisWorkbook.Sheets("Purchase_Followup").ListObjects("tbl_PurchaseOrders")
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Dim r As Long: r = TblRow(lo,"PO_ID",poID)
    If r = 0 Then Exit Sub
    Dim newRec As Long
    newRec = SafeLng(lo.DataBodyRange(r,lo.ListColumns("Qty_Received").Index).Value) + qtyRec
    lo.DataBodyRange(r,lo.ListColumns("Qty_Received").Index).Value = newRec
    Dim ord As Long: ord = SafeLng(lo.DataBodyRange(r,lo.ListColumns("Qty_Ordered").Index).Value)
    lo.DataBodyRange(r,lo.ListColumns("PO_Status").Index).Value = IIf(newRec>=ord,"Received","Partial")
    If newRec >= ord Then lo.DataBodyRange(r,lo.ListColumns("Actual_Delivery").Index).Value = Date
End Sub
"""

VBA_DASH = """\
Option Explicit

Public Sub RefreshAll()
    Application.ScreenUpdating = False
    Application.CalculateFull
    WriteAlerts
    WriteShortage
    WriteRecentTxns
    ThisWorkbook.Sheets("Dashboard").Range("dash_LastRefreshed").Value = Now()
    ColorKPIs
    Application.ScreenUpdating = True
End Sub

Private Sub WriteAlerts()
    Dim wD As Worksheet: Set wD = ThisWorkbook.Sheets("Dashboard")
    Dim lo As ListObject: Set lo = ThisWorkbook.Sheets("Inventory").ListObjects("tbl_Inventory")
    Dim SR As Long: SR = 17
    wD.Range(wD.Cells(SR,2),wD.Cells(SR+6,7)).ClearContents
    wD.Range(wD.Cells(SR,2),wD.Cells(SR+6,7)).Interior.ColorIndex = xlNone
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Dim out As Long: out = SR
    Dim i As Long
    For i = 1 To lo.ListRows.Count
        If SafeStr(lo.DataBodyRange(i,lo.ListColumns("Alert_Level").Index).Value) = "RED_ALERT" Then
            wD.Cells(out,2).Value = lo.DataBodyRange(i,lo.ListColumns("Item_ID").Index).Value
            wD.Cells(out,3).Value = lo.DataBodyRange(i,lo.ListColumns("Item_Name_HE").Index).Value
            wD.Cells(out,4).Value = lo.DataBodyRange(i,lo.ListColumns("Manufacturer").Index).Value
            wD.Cells(out,5).Value = lo.DataBodyRange(i,lo.ListColumns("Qty_On_Hand").Index).Value
            wD.Cells(out,6).Value = lo.DataBodyRange(i,lo.ListColumns("Min_Qty").Index).Value
            wD.Range(wD.Cells(out,2),wD.Cells(out,6)).Interior.Color = RGB(255,180,180)
            out = out+1: If out > SR+6 Then Exit For
        End If
    Next i
End Sub

Private Sub WriteShortage()
    Dim wD As Worksheet: Set wD = ThisWorkbook.Sheets("Dashboard")
    Dim lo As ListObject: Set lo = ThisWorkbook.Sheets("Inventory").ListObjects("tbl_Inventory")
    Dim SR As Long: SR = 26
    wD.Range(wD.Cells(SR,2),wD.Cells(SR+8,7)).ClearContents
    wD.Range(wD.Cells(SR,2),wD.Cells(SR+8,7)).Interior.ColorIndex = xlNone
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Dim out As Long: out = SR
    Dim i As Long
    For i = 1 To lo.ListRows.Count
        If SafeStr(lo.DataBodyRange(i,lo.ListColumns("Shortage_Flag").Index).Value) = "חסר" Then
            wD.Cells(out,2).Value = lo.DataBodyRange(i,lo.ListColumns("Item_ID").Index).Value
            wD.Cells(out,3).Value = lo.DataBodyRange(i,lo.ListColumns("Item_Name_HE").Index).Value
            wD.Cells(out,4).Value = lo.DataBodyRange(i,lo.ListColumns("Qty_On_Hand").Index).Value
            wD.Cells(out,5).Value = lo.DataBodyRange(i,lo.ListColumns("Min_Qty").Index).Value
            wD.Cells(out,6).Value = lo.DataBodyRange(i,lo.ListColumns("Is_Critical").Index).Value
            Dim al As String: al = SafeStr(lo.DataBodyRange(i,lo.ListColumns("Alert_Level").Index).Value)
            wD.Range(wD.Cells(out,2),wD.Cells(out,6)).Interior.Color = IIf(al="RED_ALERT",RGB(255,180,180),RGB(255,235,180))
            out = out+1: If out > SR+8 Then Exit For
        End If
    Next i
End Sub

Private Sub WriteRecentTxns()
    Dim wD As Worksheet: Set wD = ThisWorkbook.Sheets("Dashboard")
    Dim lo As ListObject: Set lo = ThisWorkbook.Sheets("Transactions").ListObjects("tbl_Transactions")
    Dim SR As Long: SR = 37
    wD.Range(wD.Cells(SR,2),wD.Cells(SR+9,9)).ClearContents
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Dim total As Long: total = lo.ListRows.Count
    Dim st As Long: st = IIf(total>10,total-9,1)
    Dim out As Long: out = SR
    Dim i As Long
    For i = total To st Step -1
        wD.Cells(out,2).Value = lo.DataBodyRange(i,lo.ListColumns("Txn_ID").Index).Value
        wD.Cells(out,3).Value = lo.DataBodyRange(i,lo.ListColumns("Txn_Date").Index).Value
        wD.Cells(out,4).Value = lo.DataBodyRange(i,lo.ListColumns("Txn_Type").Index).Value
        wD.Cells(out,5).Value = lo.DataBodyRange(i,lo.ListColumns("Item_ID").Index).Value
        wD.Cells(out,6).Value = lo.DataBodyRange(i,lo.ListColumns("Item_Name_HE").Index).Value
        wD.Cells(out,7).Value = lo.DataBodyRange(i,lo.ListColumns("Qty_Change").Index).Value
        wD.Cells(out,8).Value = lo.DataBodyRange(i,lo.ListColumns("Technician").Index).Value
        wD.Cells(out,9).Value = lo.DataBodyRange(i,lo.ListColumns("Reason").Index).Value
        out = out+1: If out > SR+9 Then Exit For
    Next i
End Sub

Private Sub ColorKPIs()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Dashboard")
    Dim crit As Long: crit = SafeLng(ws.Range("dash_CriticalZero").Value)
    With ws.Range("dash_CriticalZero")
        .Interior.Color = IIf(crit>0,RGB(220,0,0),RGB(0,176,80))
        .Font.Color = RGB(255,255,255)
    End With
    Dim sh As Long: sh = SafeLng(ws.Range("dash_ShortageCount").Value)
    With ws.Range("dash_ShortageCount")
        .Interior.Color = IIf(sh>0,RGB(255,165,0),RGB(0,176,80))
        .Font.Color = RGB(255,255,255)
    End With
End Sub
"""

VBA_REPORTS = """\
Option Explicit

Public Sub ExportShortage()
    Dim loI As ListObject
    Set loI = ThisWorkbook.Sheets("Inventory").ListObjects("tbl_Inventory")
    If loI.DataBodyRange Is Nothing Then ShowErr "אין נתוני מלאי": Exit Sub
    Dim nm As String: nm = "חוסרים_" & Format(Now,"YYYYMMDD")
    On Error Resume Next
    Dim old As Worksheet: Set old = ThisWorkbook.Sheets(nm): On Error GoTo 0
    If Not old Is Nothing Then Application.DisplayAlerts=False: old.Delete: Application.DisplayAlerts=True
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = nm
    ws.Tab.Color = RGB(220,0,0)
    ws.DisplayRightToLeft = True
    ws.Range("A1").Value = "דוח חוסרים - " & Format(Now,"DD/MM/YYYY HH:MM")
    ws.Range("A3:G3").Value = Array("מזהה","שם פריט","יצרן","מלאי","מינימום","קריטי","רמה")
    ws.Range("A3:G3").Font.Bold = True
    ws.Range("A3:G3").Interior.Color = RGB(68,114,196)
    ws.Range("A3:G3").Font.Color = RGB(255,255,255)
    Dim out As Long: out = 4
    Dim i As Long
    For i = 1 To loI.ListRows.Count
        If SafeStr(loI.DataBodyRange(i,loI.ListColumns("Shortage_Flag").Index).Value) = "חסר" Then
            ws.Cells(out,1).Value = loI.DataBodyRange(i,loI.ListColumns("Item_ID").Index).Value
            ws.Cells(out,2).Value = loI.DataBodyRange(i,loI.ListColumns("Item_Name_HE").Index).Value
            ws.Cells(out,3).Value = loI.DataBodyRange(i,loI.ListColumns("Manufacturer").Index).Value
            ws.Cells(out,4).Value = loI.DataBodyRange(i,loI.ListColumns("Qty_On_Hand").Index).Value
            ws.Cells(out,5).Value = loI.DataBodyRange(i,loI.ListColumns("Min_Qty").Index).Value
            ws.Cells(out,6).Value = loI.DataBodyRange(i,loI.ListColumns("Is_Critical").Index).Value
            ws.Cells(out,7).Value = loI.DataBodyRange(i,loI.ListColumns("Alert_Level").Index).Value
            Dim al As String: al = SafeStr(loI.DataBodyRange(i,loI.ListColumns("Alert_Level").Index).Value)
            ws.Range(ws.Cells(out,1),ws.Cells(out,7)).Interior.Color = IIf(al="RED_ALERT",RGB(255,180,180),RGB(255,235,180))
            out = out+1
        End If
    Next i
    ws.Columns("A:G").AutoFit
    ws.Activate
    ShowOK "דוח חוסרים הופק: " & (out-4) & " פריטים"
End Sub
"""

VBA_MIGRATION = """\
Option Explicit

Public Sub RunMigration()
    Dim src As String
    src = ThisWorkbook.Path & "\\" & Chr(1502) & Chr(1492) & Chr(1491) & Chr(1493) & Chr(1512) & Chr(1492) & _
          " -1 " & Chr(1504) & Chr(1497) & Chr(1492) & Chr(1493) & Chr(1500) & " " & Chr(1502) & Chr(1500) & _
          Chr(1488) & Chr(1497) & " " & Chr(1495) & Chr(1500) & Chr(1511) & Chr(1497) & " " & Chr(1495) & _
          Chr(1497) & Chr(1500) & Chr(1493) & Chr(1507) & " " & Chr(1502) & Chr(1499) & Chr(1513) & Chr(1493) & Chr(1512) & ".xlsx"
    If Dir(src) = "" Then
        src = Application.GetOpenFilename("Excel Files (*.xlsx),*.xlsx",,"בחר קובץ מקור")
        If src = "False" Then Exit Sub
    End If
    If MsgBox("ייבא נתונים מ:" & vbCrLf & src & vbCrLf & "האם להמשיך?",vbYesNo+vbQuestion,"ייבוא") = vbNo Then Exit Sub
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Dim wb As Workbook: Set wb = Workbooks.Open(src,ReadOnly:=True)
    Dim cnt As Long: cnt = ImportSheet6(wb)
    SeedOpeningTxns
    wb.Close False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    modDashboard.RefreshAll
    ShowOK "ייבוא הסתיים. נוספו " & cnt & " פריטים חדשים."
End Sub

Private Function ImportSheet6(wb As Workbook) As Long
    Dim ws As Worksheet
    On Error Resume Next: Set ws = wb.Sheets(6): On Error GoTo 0
    If ws Is Nothing Then ImportSheet6 = 0: Exit Function
    Dim cnt As Long: cnt = 0
    Dim r As Long
    For r = 5 To 400
        Dim nH As String, mfg As String, mdl As String
        nH  = Trim(SafeStr(ws.Cells(r,3).Value))
        mfg = Trim(SafeStr(ws.Cells(r,9).Value))
        mdl = Trim(SafeStr(ws.Cells(r,6).Value))
        If nH = "" And r > 15 Then Exit For
        If nH = "" Or (mfg = "" And mdl = "") Then GoTo Nxt
        If ItemByMfgModel(mfg,mdl) <> "" Then GoTo Nxt
        Dim it As String: it = Trim(SafeStr(ws.Cells(r,4).Value))
        If UCase(it) <> "CON" Then it = "INS"
        Dim prc As Currency: prc = SafeCur(ws.Cells(r,11).Value)
        Dim qH As Long: If IsNumeric(ws.Cells(r,12).Value) Then qH = CLng(ws.Cells(r,12).Value)
        Dim mQ As Long: If IsNumeric(ws.Cells(r,13).Value) Then mQ = CLng(ws.Cells(r,13).Value)
        Dim newID As String
        newID = AddItem(nH,Trim(SafeStr(ws.Cells(r,5).Value)),"Other","",mfg,mdl, _
            Trim(SafeStr(ws.Cells(r,10).Value)),it,"יח'",prc,0, _
            Trim(SafeStr(ws.Cells(r,14).Value)), _
            Trim(SafeStr(ws.Cells(r,7).Value)),Trim(SafeStr(ws.Cells(r,8).Value)), _
            Trim(SafeStr(ws.Cells(r,15).Value)),"No",0,0,"No","Active", _
            Trim(SafeStr(ws.Cells(r,16).Value)),mQ,0)
        If newID <> "" And qH > 0 Then
            Dim loI As ListObject
            Set loI = ThisWorkbook.Sheets("Inventory").ListObjects("tbl_Inventory")
            Dim ri As Long: ri = TblRow(loI,"Item_ID",newID)
            If ri > 0 Then loI.DataBodyRange(ri,loI.ListColumns("Qty_On_Hand").Index).Value = qH
        End If
        If newID <> "" Then cnt = cnt+1
Nxt:
    Next r
    ImportSheet6 = cnt
End Function

Private Sub SeedOpeningTxns()
    Dim loI As ListObject
    Set loI = ThisWorkbook.Sheets("Inventory").ListObjects("tbl_Inventory")
    If loI.DataBodyRange Is Nothing Then Exit Sub
    Dim i As Long
    For i = 1 To loI.ListRows.Count
        Dim id As String: id = SafeStr(loI.DataBodyRange(i,loI.ListColumns("Item_ID").Index).Value)
        Dim q  As Long:   q  = SafeLng(loI.DataBodyRange(i,loI.ListColumns("Qty_On_Hand").Index).Value)
        If id <> "" And q > 0 Then
            LogTxn "INITIAL",id,q,0,q,SafeCur(GetItemFld(id,"Unit_Price_ILS")),"פתיחת מערכת","","","Migration",""
        End If
    Next i
End Sub
"""

VBA_MAIN = """\
Option Explicit
Public Sub OpenMenu():    frmMain.Show:         End Sub
Public Sub BtnAddItem():  frmAddItem.Show:      End Sub
Public Sub BtnStockIn():  frmStockIn.Show:      End Sub
Public Sub BtnStockOut(): frmStockOut.Show:     End Sub
Public Sub BtnAdjust():   frmAdjust.Show:       End Sub
Public Sub BtnSearch():   frmSearch.Show:       End Sub
Public Sub BtnRefresh():  modDashboard.RefreshAll: ShowOK "לוח הבקרה עודכן": End Sub
Public Sub BtnOrders():   modReports.ExportShortage: End Sub
"""

VBA_THISWORKBOOK = """\
Private Sub Workbook_Open()
    Application.ScreenUpdating = False
    modDashboard.RefreshAll
    Application.ScreenUpdating = True
    ThisWorkbook.Sheets("Dashboard").Activate
End Sub
"""

# ── UserForm code ────────────────────────────────────────────────────────────
CODE_FRMMAIN = """\
Option Explicit
Private Sub UserForm_Initialize()
    Me.Width=310: Me.Height=385
    On Error Resume Next: Me.RightToLeft = True: On Error GoTo 0
    ApplyRTLAlignment Me
End Sub
Private Sub cmdAddItem_Click():  Me.Hide: frmAddItem.Show:  Me.Show: End Sub
Private Sub cmdStockIn_Click():  Me.Hide: frmStockIn.Show:  Me.Show: End Sub
Private Sub cmdStockOut_Click(): Me.Hide: frmStockOut.Show: Me.Show: End Sub
Private Sub cmdAdjust_Click():   Me.Hide: frmAdjust.Show:   Me.Show: End Sub
Private Sub cmdSearch_Click():   Me.Hide: frmSearch.Show:   Me.Show: End Sub
Private Sub cmdRefresh_Click():  modDashboard.RefreshAll: ShowOK "עודכן": End Sub
Private Sub cmdOrders_Click():   modReports.ExportShortage: End Sub
Private Sub cmdClose_Click():    Unload Me: End Sub
"""

CODE_FRMSTOCKOUT = """\
Option Explicit
Dim selID As String, curQ As Long, minQ As Long, isCrit As Boolean

Private Sub UserForm_Initialize()
    Me.Width=492: Me.Height=500
    On Error Resume Next: Me.RightToLeft = True: On Error GoTo 0
    ApplyRTLAlignment Me
    txtTechnician.Value = Environ("USERNAME")
    lblWarn.Visible = False: lblCrit.Visible = False
    selID = "": curQ = 0: minQ = 0: isCrit = False
End Sub

Private Sub txtSearch_Change()
    lstRes.Clear
    If Len(txtSearch.Value) < 2 Then Exit Sub
    Dim lo As ListObject
    Set lo = ThisWorkbook.Sheets("Items_Master").ListObjects("tbl_Items")
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Dim s As String: s = LCase(Trim(txtSearch.Value))
    Dim i As Long, c As Long: c = 0
    For i = 1 To lo.ListRows.Count
        Dim nH As String: nH = LCase(SafeStr(lo.DataBodyRange(i,lo.ListColumns("Item_Name_HE").Index).Value))
        Dim mf As String: mf = LCase(SafeStr(lo.DataBodyRange(i,lo.ListColumns("Manufacturer").Index).Value))
        Dim md As String: md = LCase(SafeStr(lo.DataBodyRange(i,lo.ListColumns("Model").Index).Value))
        Dim pn As String: pn = LCase(SafeStr(lo.DataBodyRange(i,lo.ListColumns("Part_Number").Index).Value))
        Dim ct As String: ct = LCase(SafeStr(lo.DataBodyRange(i,lo.ListColumns("Category").Index).Value))
        Dim iid2 As String: iid2 = LCase(SafeStr(lo.DataBodyRange(i,lo.ListColumns("Item_ID").Index).Value))
        If InStr(nH,s)>0 Or InStr(mf,s)>0 Or InStr(md,s)>0 Or InStr(pn,s)>0 Or InStr(ct,s)>0 Or InStr(iid2,s)>0 Then
            Dim iid As String: iid = SafeStr(lo.DataBodyRange(i,lo.ListColumns("Item_ID").Index).Value)
            lstRes.AddItem SafeStr(lo.DataBodyRange(i,lo.ListColumns("Item_Name_HE").Index).Value)
            lstRes.List(lstRes.ListCount-1,1) = SafeStr(lo.DataBodyRange(i,lo.ListColumns("Manufacturer").Index).Value)
            lstRes.List(lstRes.ListCount-1,2) = SafeStr(lo.DataBodyRange(i,lo.ListColumns("Model").Index).Value)
            lstRes.List(lstRes.ListCount-1,3) = CStr(modInventory.GetQty(iid))
            lstRes.List(lstRes.ListCount-1,4) = iid
            c = c+1: If c >= 50 Then Exit For
        End If
    Next i
End Sub

Private Sub lstRes_Click()
    If lstRes.ListIndex < 0 Then Exit Sub
    selID = lstRes.List(lstRes.ListIndex,4)
    curQ  = modInventory.GetQty(selID)
    Dim loM As ListObject: Set loM = ThisWorkbook.Sheets("Min_Stock").ListObjects("tbl_MinStock")
    Dim rm As Long: rm = TblRow(loM,"Item_ID",selID)
    minQ = IIf(rm>0,SafeLng(loM.DataBodyRange(rm,loM.ListColumns("Min_Qty").Index).Value),0)
    isCrit = (SafeStr(GetItemFld(selID,"Is_Critical")) = "Yes")
    lblSel.Caption = lstRes.List(lstRes.ListIndex,0) & " | " & lstRes.List(lstRes.ListIndex,1)
    lblCurQ.Caption = "מלאי נוכחי: " & curQ
    lblMinQ.Caption = "מינימום: " & minQ
    lblCurQ.ForeColor = IIf(curQ=0,RGB(200,0,0),IIf(curQ<minQ,RGB(200,100,0),RGB(0,100,0)))
    txtQty.Value = "": lblWarn.Visible = False: lblCrit.Visible = False
End Sub

Private Sub txtQty_Change()
    lblWarn.Visible = False: lblCrit.Visible = False: cmdSave.Enabled = True
    If selID = "" Or Not IsNumeric(txtQty.Value) Then Exit Sub
    Dim q As Long: q = SafeLng(txtQty.Value)
    If q <= 0 Then cmdSave.Enabled = False: Exit Sub
    If q > curQ Then
        lblWarn.Caption = "כמות (" & q & ") גדולה מהמלאי (" & curQ & ")!"
        lblWarn.ForeColor = RGB(200,0,0): lblWarn.Visible = True: cmdSave.Enabled = False: Exit Sub
    End If
    If curQ-q = 0 And isCrit Then
        lblCrit.Caption = "אזהרה! פריט קריטי (שבר, משבית מערכת) יגיע לאפס!"
        lblCrit.ForeColor = RGB(200,0,0): lblCrit.Visible = True
    ElseIf curQ-q < minQ Then
        lblWarn.Caption = "מלאי יהיה מתחת למינימום (" & (curQ-q) & " < " & minQ & ")"
        lblWarn.ForeColor = RGB(200,100,0): lblWarn.Visible = True
    End If
End Sub

Private Sub cmdSave_Click()
    If selID = "" Then ShowErr "יש לבחור פריט": Exit Sub
    If Not IsNumeric(txtQty.Value) Or SafeLng(txtQty.Value)<=0 Then ShowErr "כמות לא תקינה": Exit Sub
    If Len(Trim(txtTechnician.Value))=0 Then ShowErr "שם מבצע הוא חובה": Exit Sub
    modInventory.StockOut selID,SafeLng(txtQty.Value),Trim(txtReason.Value),Trim(txtWO.Value),Trim(txtTechnician.Value),""
    ShowOK "הוצאה בוצעה!"
    txtSearch.Value="": lstRes.Clear: txtQty.Value=""
    lblSel.Caption="": lblCurQ.Caption="מלאי: -": lblMinQ.Caption="מינימום: -"
    lblWarn.Visible=False: lblCrit.Visible=False: selID="": curQ=0
End Sub
Private Sub cmdCancel_Click(): Unload Me: End Sub
"""

CODE_FRMSTOCKIN = """\
Option Explicit
Dim selID As String

Private Sub UserForm_Initialize()
    Me.Width=492: Me.Height=485
    On Error Resume Next: Me.RightToLeft = True: On Error GoTo 0
    ApplyRTLAlignment Me
    txtTechnician.Value = Environ("USERNAME")
    selID = ""
End Sub

Private Sub txtSearch_Change()
    lstRes.Clear
    If Len(txtSearch.Value) < 2 Then Exit Sub
    Dim lo As ListObject
    Set lo = ThisWorkbook.Sheets("Items_Master").ListObjects("tbl_Items")
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Dim s As String: s = LCase(Trim(txtSearch.Value))
    Dim i As Long, c As Long: c = 0
    For i = 1 To lo.ListRows.Count
        Dim nH As String: nH = LCase(SafeStr(lo.DataBodyRange(i,lo.ListColumns("Item_Name_HE").Index).Value))
        Dim mf As String: mf = LCase(SafeStr(lo.DataBodyRange(i,lo.ListColumns("Manufacturer").Index).Value))
        Dim md As String: md = LCase(SafeStr(lo.DataBodyRange(i,lo.ListColumns("Model").Index).Value))
        Dim pn As String: pn = LCase(SafeStr(lo.DataBodyRange(i,lo.ListColumns("Part_Number").Index).Value))
        Dim ct As String: ct = LCase(SafeStr(lo.DataBodyRange(i,lo.ListColumns("Category").Index).Value))
        Dim iid2 As String: iid2 = LCase(SafeStr(lo.DataBodyRange(i,lo.ListColumns("Item_ID").Index).Value))
        If InStr(nH,s)>0 Or InStr(mf,s)>0 Or InStr(md,s)>0 Or InStr(pn,s)>0 Or InStr(ct,s)>0 Or InStr(iid2,s)>0 Then
            Dim iid As String: iid = SafeStr(lo.DataBodyRange(i,lo.ListColumns("Item_ID").Index).Value)
            lstRes.AddItem SafeStr(lo.DataBodyRange(i,lo.ListColumns("Item_Name_HE").Index).Value)
            lstRes.List(lstRes.ListCount-1,1) = SafeStr(lo.DataBodyRange(i,lo.ListColumns("Manufacturer").Index).Value)
            lstRes.List(lstRes.ListCount-1,2) = CStr(modInventory.GetQty(iid))
            lstRes.List(lstRes.ListCount-1,3) = iid
            c = c+1: If c >= 50 Then Exit For
        End If
    Next i
End Sub

Private Sub lstRes_Click()
    If lstRes.ListIndex < 0 Then Exit Sub
    selID = lstRes.List(lstRes.ListIndex,3)
    lblSel.Caption  = lstRes.List(lstRes.ListIndex,0) & " | " & lstRes.List(lstRes.ListIndex,1)
    lblCurQ.Caption = "מלאי נוכחי: " & modInventory.GetQty(selID)
    txtPrice.Value  = SafeStr(GetItemFld(selID,"Unit_Price_ILS"))
End Sub

Private Sub cmdSave_Click()
    If selID = "" Then ShowErr "יש לבחור פריט": Exit Sub
    If Not IsNumeric(txtQty.Value) Or SafeLng(txtQty.Value)<=0 Then ShowErr "כמות לא תקינה": Exit Sub
    If Len(Trim(txtTechnician.Value))=0 Then ShowErr "שם מקבל הוא חובה": Exit Sub
    modInventory.StockIn selID,SafeLng(txtQty.Value),SafeCur(txtPrice.Value),Trim(txtPO.Value),Trim(txtReason.Value),Trim(txtTechnician.Value)
    ShowOK "קליטה בוצעה!"
    txtSearch.Value="": lstRes.Clear: txtQty.Value=""
    lblSel.Caption="": lblCurQ.Caption="מלאי: -": selID=""
End Sub
Private Sub cmdCancel_Click(): Unload Me: End Sub
"""

CODE_FRMADDITEM = """\
Option Explicit

Private Sub UserForm_Initialize()
    Me.Width=432: Me.Height=510
    On Error Resume Next: Me.RightToLeft = True: On Error GoTo 0
    ApplyRTLAlignment Me
    Dim lo As ListObject
    Dim i As Long
    Set lo = ThisWorkbook.Sheets("Lists").ListObjects("tbl_Categories")
    If Not lo.DataBodyRange Is Nothing Then
        For i = 1 To lo.ListRows.Count: cboCategory.AddItem SafeStr(lo.DataBodyRange(i,1).Value): Next i
    End If
    Set lo = ThisWorkbook.Sheets("Lists").ListObjects("tbl_Manufacturers")
    If Not lo.DataBodyRange Is Nothing Then
        For i = 1 To lo.ListRows.Count: cboMfg.AddItem SafeStr(lo.DataBodyRange(i,1).Value): Next i
    End If
    Set lo = ThisWorkbook.Sheets("Lists").ListObjects("tbl_Units")
    If Not lo.DataBodyRange Is Nothing Then
        For i = 1 To lo.ListRows.Count: cboUnit.AddItem SafeStr(lo.DataBodyRange(i,1).Value): Next i
    End If
    Set lo = ThisWorkbook.Sheets("Lists").ListObjects("tbl_Locations")
    If Not lo.DataBodyRange Is Nothing Then
        For i = 1 To lo.ListRows.Count: cboLoc.AddItem SafeStr(lo.DataBodyRange(i,1).Value): Next i
    End If
    cboType.AddItem "INS": cboType.AddItem "CON"
    chkCritical.Value = False
    Dim seq As Long: seq = SafeLng(ThisWorkbook.Names("cfg_NextItemSeq").RefersToRange.Value)
    lblPreview.Caption = "מזהה שיוקצה: ITM-" & Year(Now) & "-" & Format(seq,"0000")
End Sub

Private Sub cmdSave_Click()
    If Len(Trim(txtNameHE.Value))=0 Then ShowErr "שם עברי הוא חובה":   Exit Sub
    If Len(Trim(txtModel.Value))=0  Then ShowErr "דגם הוא חובה":        Exit Sub
    If Len(cboCategory.Value)=0    Then ShowErr "קטגוריה היא חובה":     Exit Sub
    If Len(cboMfg.Value)=0         Then ShowErr "יצרן הוא חובה":        Exit Sub
    Dim newID As String
    newID = AddItem( _
        Trim(txtNameHE.Value),Trim(txtNameEN.Value), _
        cboCategory.Value,"", _
        cboMfg.Value,Trim(txtModel.Value),Trim(txtPN.Value), _
        IIf(cboType.Value="","INS",cboType.Value), _
        IIf(cboUnit.Value="","יח'",cboUnit.Value), _
        SafeCur(txtPrice.Value),0,cboLoc.Value,"","","", _
        IIf(chkCritical.Value,"Yes","No"),0,0,"No","Active", _
        Trim(txtNotes.Value),SafeLng(txtMinQty.Value),0)
    If newID = "" Then Exit Sub
    modDashboard.RefreshAll
    ShowOK "פריט נוסף: " & newID
    If MsgBox("להוסיף פריט נוסף?",vbYesNo+vbQuestion,"הוספה") = vbYes Then
        txtNameHE.Value="": txtNameEN.Value="": txtModel.Value=""
        txtPN.Value="": txtPrice.Value="": txtMinQty.Value="0": txtNotes.Value=""
        chkCritical.Value=False
        Dim seq As Long: seq = SafeLng(ThisWorkbook.Names("cfg_NextItemSeq").RefersToRange.Value)
        lblPreview.Caption = "מזהה שיוקצה: ITM-" & Year(Now) & "-" & Format(seq,"0000")
    Else: Unload Me
    End If
End Sub
Private Sub cmdCancel_Click(): Unload Me: End Sub
"""

CODE_FRMADJUST = """\
Option Explicit
Dim selID As String

Private Sub UserForm_Initialize()
    Me.Width=402: Me.Height=310
    On Error Resume Next: Me.RightToLeft = True: On Error GoTo 0
    ApplyRTLAlignment Me
    txtTechnician.Value = Environ("USERNAME")
    selID = ""
End Sub

Private Sub txtItemID_Change()
    selID = Trim(txtItemID.Value)
    If ItemByID(selID) Then
        lblName.Caption = SafeStr(GetItemFld(selID,"Item_Name_HE"))
        lblCurQ.Caption = "מלאי נוכחי: " & modInventory.GetQty(selID)
    Else: lblName.Caption = "": lblCurQ.Caption = ""
    End If
End Sub

Private Sub txtNewQty_Change()
    If selID = "" Or Not IsNumeric(txtNewQty.Value) Then lblDelta.Caption = "": Exit Sub
    Dim delta As Long: delta = SafeLng(txtNewQty.Value) - modInventory.GetQty(selID)
    lblDelta.Caption = IIf(delta>=0,"+" & delta,CStr(delta))
    lblDelta.ForeColor = IIf(delta>=0,RGB(0,150,0),RGB(200,0,0))
End Sub

Private Sub cmdSave_Click()
    If Not ItemByID(selID) Then ShowErr "פריט לא נמצא": Exit Sub
    If Not IsNumeric(txtNewQty.Value) Or SafeLng(txtNewQty.Value)<0 Then ShowErr "כמות לא תקינה": Exit Sub
    If Len(Trim(txtReason.Value))=0 Then ShowErr "סיבה היא חובה לתיקון מלאי": Exit Sub
    If Len(Trim(txtTechnician.Value))=0 Then ShowErr "שם מבצע הוא חובה": Exit Sub
    modInventory.StockAdjust selID,SafeLng(txtNewQty.Value),Trim(txtReason.Value),Trim(txtTechnician.Value)
    ShowOK "תיקון מלאי בוצע!"
    Unload Me
End Sub
Private Sub cmdCancel_Click(): Unload Me: End Sub
"""

CODE_FRMSEARCH = """\
Option Explicit

Private Sub UserForm_Initialize()
    Me.Width=652: Me.Height=478
    On Error Resume Next: Me.RightToLeft = True: On Error GoTo 0
    ApplyRTLAlignment Me
End Sub
Private Sub cmdSearch_Click()
    lstRes.Clear
    Dim lo As ListObject
    Set lo = ThisWorkbook.Sheets("Items_Master").ListObjects("tbl_Items")
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Dim s As String: s = LCase(Trim(txtSearch.Value))
    Dim i As Long, c As Long: c = 0
    For i = 1 To lo.ListRows.Count
        Dim nH As String: nH = LCase(SafeStr(lo.DataBodyRange(i,lo.ListColumns("Item_Name_HE").Index).Value))
        Dim mf As String: mf = LCase(SafeStr(lo.DataBodyRange(i,lo.ListColumns("Manufacturer").Index).Value))
        Dim md As String: md = LCase(SafeStr(lo.DataBodyRange(i,lo.ListColumns("Model").Index).Value))
        Dim ct As String: ct = LCase(SafeStr(lo.DataBodyRange(i,lo.ListColumns("Category").Index).Value))
        Dim pn As String: pn = LCase(SafeStr(lo.DataBodyRange(i,lo.ListColumns("Part_Number").Index).Value))
        Dim loc As String: loc = LCase(SafeStr(lo.DataBodyRange(i,lo.ListColumns("Storage_Location").Index).Value))
        Dim iid3 As String: iid3 = LCase(SafeStr(lo.DataBodyRange(i,lo.ListColumns("Item_ID").Index).Value))
        If s="" Or InStr(nH,s)>0 Or InStr(mf,s)>0 Or InStr(md,s)>0 Or InStr(ct,s)>0 Or InStr(pn,s)>0 Or InStr(loc,s)>0 Or InStr(iid3,s)>0 Then
            Dim iid As String: iid = SafeStr(lo.DataBodyRange(i,lo.ListColumns("Item_ID").Index).Value)
            lstRes.AddItem iid
            lstRes.List(lstRes.ListCount-1,1) = SafeStr(lo.DataBodyRange(i,lo.ListColumns("Item_Name_HE").Index).Value)
            lstRes.List(lstRes.ListCount-1,2) = SafeStr(lo.DataBodyRange(i,lo.ListColumns("Manufacturer").Index).Value)
            lstRes.List(lstRes.ListCount-1,3) = SafeStr(lo.DataBodyRange(i,lo.ListColumns("Model").Index).Value)
            lstRes.List(lstRes.ListCount-1,4) = CStr(modInventory.GetQty(iid))
            lstRes.List(lstRes.ListCount-1,5) = SafeStr(lo.DataBodyRange(i,lo.ListColumns("Is_Critical").Index).Value)
            c = c+1
        End If
    Next i
    lblCount.Caption = "נמצאו " & c & " פריטים"
End Sub

Private Sub lstRes_Click()
    If lstRes.ListIndex < 0 Then Exit Sub
    Dim iid As String: iid = lstRes.List(lstRes.ListIndex,0)
    lblDID.Caption   = iid
    lblDName.Caption = SafeStr(GetItemFld(iid,"Item_Name_HE"))
    lblDMfg.Caption  = SafeStr(GetItemFld(iid,"Manufacturer"))
    lblDModel.Caption= SafeStr(GetItemFld(iid,"Model"))
    Dim q As Long: q = modInventory.GetQty(iid)
    lblDQty.Caption  = CStr(q)
    lblDQty.ForeColor = IIf(q=0,RGB(200,0,0),RGB(0,100,0))
    lblDLoc.Caption  = SafeStr(GetItemFld(iid,"Storage_Location"))
    lblDCrit.Caption = SafeStr(GetItemFld(iid,"Is_Critical"))
End Sub

Private Sub cmdStockOut_Click()
    Me.Hide: frmStockOut.Show: Me.Show
End Sub
Private Sub cmdStockIn_Click()
    Me.Hide: frmStockIn.Show: Me.Show
End Sub
Private Sub cmdClose_Click(): Unload Me: End Sub
"""

# ════════════════════════════════════════════════════════════════════════════
#  BUILDER FUNCTIONS
# ════════════════════════════════════════════════════════════════════════════

def enable_vba_access():
    for ver in ["16.0","15.0","14.0","17.0"]:
        try:
            k = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                               f"Software\\Microsoft\\Office\\{ver}\\Excel\\Security",
                               0, winreg.KEY_SET_VALUE)
            winreg.SetValueEx(k,"AccessVBOM",0,winreg.REG_DWORD,1)
            winreg.CloseKey(k)
            print(f"  VBA access enabled (Office {ver})")
            return
        except: pass


def get_ws(wb, name):
    """Get sheet by name via xlwings."""
    return wb.sheets[name]


def ws_api(wb, name):
    """Get the underlying COM object for a sheet."""
    return wb.sheets[name].api


def make_table(wb, sheet_name, headers, table_name, style="TableStyleMedium2"):
    """Write headers to row 1 and create a ListObject table."""
    ws = wb.sheets[sheet_name]
    ws.range("A1").value = headers          # writes as a row
    ncols = len(headers)
    rng_addr = f"A1:{chr(64+ncols)}1"
    ws_a = ws.api
    rng = ws_a.Range(rng_addr)
    lo = ws_a.ListObjects.Add(1, rng, None, 1)   # xlSrcRange, rng, link, xlYes
    lo.Name = table_name
    lo.TableStyle = style
    return lo


def setup_sheets(wb):
    """Delete default sheets, create all 11 with correct names, RTL, tab colors."""
    app = wb.app

    needed = [n for n, _ in SHEET_DEFS]
    existing = [s.name for s in wb.sheets]

    # Add missing sheets
    for name in needed:
        if name not in existing:
            wb.sheets.add(name, after=wb.sheets[-1])

    # Delete sheets not in our list
    app.api.DisplayAlerts = False
    for s in list(wb.sheets):
        if s.name not in needed:
            try: s.delete()
            except: pass
    app.api.DisplayAlerts = True

    # Reorder and style
    for idx, (name, color) in enumerate(SHEET_DEFS):
        s = wb.sheets[name]
        # Move to correct position
        try:
            s.api.Move(Before=wb.sheets[idx].api)
        except: pass
        s.api.Tab.Color = color
        s.api.DisplayRightToLeft = True
    print(f"  {len(SHEET_DEFS)} sheets created and styled")


def setup_lists(wb):
    ws  = wb.sheets["Lists"]
    wsa = ws.api

    def put_list(data, col_letter, table_name):
        ws.range(f"{col_letter}1").value = [table_name.replace("tbl_","")]
        for i, v in enumerate(data, 2):
            ws.range(f"{col_letter}{i}").value = v
        n = len(data)+1
        rng = wsa.Range(f"{col_letter}1:{col_letter}{n}")
        lo  = wsa.ListObjects.Add(1, rng, None, 1)
        lo.Name = table_name

    put_list(CATEGORIES,   "A", "tbl_Categories")
    put_list(MANUFACTURERS,"C", "tbl_Manufacturers")
    put_list(LOCATIONS,    "E", "tbl_Locations")
    put_list(TXN_TYPES,    "G", "tbl_TxnTypes")
    put_list(UNITS,        "I", "tbl_Units")
    put_list(ITEM_STATUS,  "K", "tbl_ItemStatus")
    put_list(PO_STATUS,    "M", "tbl_POStatus")
    put_list(YESNO,        "O", "tbl_YesNo")
    put_list(ITEM_TYPES,   "Q", "tbl_ItemTypes")
    put_list(REASONS,      "S", "tbl_Reasons")
    wsa.Columns("A:T").AutoFit()
    print("  Lists ready")


def setup_settings(wb):
    ws = wb.sheets["Settings"]
    data = [
        ("A2","שם המפעל"),          ("B2","(הכנס שם מפעל)"),
        ("A3","מחלקה"),              ("B3","מכשור ובקרה"),
        ("A6","Item ID Prefix"),     ("B6","ITM"),
        ("A7","Txn ID Prefix"),      ("B7","TXN"),
        ("A8","PO ID Prefix"),       ("B8","PO"),
        ("A16","cfg_NextItemSeq"),   ("B16",1),
        ("A17","cfg_NextTxnSeq"),    ("B17",1),
        ("A18","cfg_NextPOSeq"),     ("B18",1),
        ("A19","cfg_NextSupplierSeq"),("B19",1),
    ]
    for cell, val in data:
        ws.range(cell).value = val
    ws.api.Columns("A:B").AutoFit()
    print("  Settings ready")


def setup_named_ranges(wb):
    pairs = {
        "cfg_NextItemSeq":     "=Settings!$B$16",
        "cfg_NextTxnSeq":      "=Settings!$B$17",
        "cfg_NextPOSeq":       "=Settings!$B$18",
        "cfg_NextSupplierSeq": "=Settings!$B$19",
        "dash_TotalItems":     "=Dashboard!$B$6",
        "dash_ActiveItems":    "=Dashboard!$B$7",
        "dash_TotalUnits":     "=Dashboard!$B$8",
        "dash_StockValue":     "=Dashboard!$B$9",
        "dash_ShortageCount":  "=Dashboard!$B$10",
        "dash_CriticalZero":   "=Dashboard!$B$11",
        "dash_OpenPOs":        "=Dashboard!$B$12",
        "dash_MonthTxns":      "=Dashboard!$B$13",
        "dash_LastRefreshed":  "=Dashboard!$G$3",
    }
    for nm, formula in pairs.items():
        try: wb.api.Names.Add(nm, formula)
        except: pass
    print("  Named ranges created")


def setup_data_sheets(wb):
    make_table(wb, "Items_Master",      H_ITEMS, "tbl_Items",          "TableStyleMedium2")
    make_table(wb, "Inventory",         H_INV,   "tbl_Inventory",      "TableStyleMedium4")
    make_table(wb, "Min_Stock",         H_MIN,   "tbl_MinStock",       "TableStyleMedium4")
    make_table(wb, "Transactions",      H_TXN,   "tbl_Transactions",   "TableStyleMedium15")
    make_table(wb, "Suppliers",         H_SUP,   "tbl_Suppliers",      "TableStyleMedium2")
    make_table(wb, "Purchase_Followup", H_PO,    "tbl_PurchaseOrders", "TableStyleMedium2")
    make_table(wb, "Assets_Link",       H_ASS,   "tbl_AssetsLink",     "TableStyleMedium18")
    make_table(wb, "Archive",           H_ARC,   "tbl_Archive",        "TableStyleLight11")

    # Column widths for key sheets
    for sname, col_widths in [
        ("Items_Master",  {1:18, 2:35, 6:22, 7:22}),
        ("Inventory",     {1:18, 2:35, 4:22}),
        ("Transactions",  {1:22, 2:18, 3:10, 5:30}),
    ]:
        wsa = wb.sheets[sname].api
        for col, w in col_widths.items():
            wsa.Columns(col).ColumnWidth = w
    print("  Data sheets ready")


def setup_dashboard(wb):
    ws  = wb.sheets["Dashboard"]
    wsa = ws.api
    app = wb.app.api

    def write(cell, val, bold=False, size=11, fcolor=None, bg=None, halign=None):
        r = wsa.Range(cell)
        r.Value = val
        if bold:   r.Font.Bold = True
        if size != 11: r.Font.Size = size
        if fcolor: r.Font.Color = fcolor
        if bg:     r.Interior.Color = bg
        if halign: r.HorizontalAlignment = halign

    wsa.Rows(1).RowHeight = 38
    wsa.Range("A1:J1").Merge()
    write("A1", "ניהול מלאי חלקי חילוף – מכשור ובקרה",
          bold=True, size=16, fcolor=rgb(255,255,255), bg=rgb(68,114,196), halign=-4108)

    wsa.Range("A2:J2").Merge()
    write("A2", "מחלקת מכשור, אוטומציה ובקרה", size=11, fcolor=rgb(68,114,196), halign=-4108)

    wsa.Range("G3").NumberFormat = "DD/MM/YYYY HH:MM"

    # KPI section
    wsa.Range("A5:D5").Merge()
    write("A5","▌ מדדי מפתח",bold=True,size=12,fcolor=rgb(68,114,196))

    kpis = [
        ("A6","B6","סה\"כ פריטים במאגר",    '=COUNTA(tbl_Items[Item_ID])'),
        ("A7","B7","פריטים פעילים",           '=COUNTIF(tbl_Items[Status],"Active")'),
        ("A8","B8","סה\"כ יחידות במלאי",      '=SUM(tbl_Inventory[Qty_On_Hand])'),
        ("A9","B9","שווי מלאי כולל ₪",        '=SUMPRODUCT(tbl_Inventory[Qty_On_Hand]*tbl_Inventory[Unit_Price_ILS])'),
        ("A10","B10","פריטים בחסר",           '=COUNTIF(tbl_Inventory[Shortage_Flag],"חסר")'),
        ("A11","B11","פריטים קריטיים בחסר",   '=COUNTIF(tbl_Inventory[Alert_Level],"RED_ALERT")'),
        ("A12","B12","הזמנות פתוחות",         '=COUNTIF(tbl_PurchaseOrders[PO_Status],"Pending")+COUNTIF(tbl_PurchaseOrders[PO_Status],"Partial")'),
        ("A13","B13","תנועות החודש",           '=COUNTIFS(tbl_Transactions[Txn_Date],">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),tbl_Transactions[Txn_Date],"<="&TODAY())'),
    ]
    for la, lb, lbl, formula in kpis:
        wsa.Range(la).Value = lbl
        wsa.Range(la).HorizontalAlignment = -4152  # right
        wsa.Range(lb).Formula = formula
        wsa.Range(lb).Font.Bold = True
        wsa.Range(lb).Font.Size = 14
        wsa.Range(lb).HorizontalAlignment = -4108  # center

    wsa.Columns("A").ColumnWidth = 32
    wsa.Columns("B").ColumnWidth = 18

    # Action buttons column E
    # Single prominent "פתח" button spanning E6:H13
    r = wsa.Range("E6:H13")
    btn = wsa.Buttons().Add(r.Left + 4, r.Top + 4, r.Width - 8, r.Height - 8)
    btn.Caption = "פתח"
    btn.OnAction = "OpenMenu"
    btn.Font.Size = 18
    btn.Font.Bold = True

    # Section headers - compact layout (fits one screen)
    wsa.Range("A15:J15").Merge()
    write("A15","▌ פריטים קריטיים בחסר (שבר, משבית מערכת)",bold=True,size=12,fcolor=rgb(255,255,255),bg=rgb(192,0,0))
    wsa.Range("B16:G16").Value = [["מזהה פריט","שם פריט","יצרן","מלאי","מינימום",""]]
    wsa.Range("B16:G16").Font.Bold = True
    wsa.Range("B16:G16").Interior.Color = rgb(192,0,0)
    wsa.Range("B16:G16").Font.Color = rgb(255,255,255)

    wsa.Range("A24:J24").Merge()
    write("A24","▌ כל הפריטים בחסר",bold=True,size=12,fcolor=rgb(255,255,255),bg=rgb(197,90,17))
    wsa.Range("B25:G25").Value = [["מזהה","שם פריט","מלאי","מינימום","קריטי",""]]
    wsa.Range("B25:G25").Font.Bold = True
    wsa.Range("B25:G25").Interior.Color = rgb(197,90,17)
    wsa.Range("B25:G25").Font.Color = rgb(255,255,255)

    wsa.Range("A35:J35").Merge()
    write("A35","▌ תנועות אחרונות",bold=True,size=12,fcolor=rgb(255,255,255),bg=rgb(68,114,196))
    wsa.Range("B36:I36").Value = [["מזהה תנועה","תאריך","סוג","מזהה","שם פריט","כמות","טכנאי","סיבה"]]
    wsa.Range("B36:I36").Font.Bold = True
    wsa.Range("B36:I36").Interior.Color = rgb(68,114,196)
    wsa.Range("B36:I36").Font.Color = rgb(255,255,255)
    print("  Dashboard ready")


def add_module(vb, name, code):
    try:
        mod = vb.VBComponents.Add(1)  # vbext_ct_StdModule
        mod.Name = name
        mod.CodeModule.AddFromString(code)
        print(f"  Module: {name}")
    except Exception as e:
        print(f"  WARNING module {name}: {e}")


def add_form(vb, name, caption, width, height, code, controls):
    try:
        frm = vb.VBComponents.Add(3)
        frm.Name = name
        d = frm.Designer
        # Set Caption/dimensions AFTER controls (setting before causes COM deadlock)
        for prog_id, ctrl_name, cap, left, top, w, h, extra in controls:
            try:
                ctrl = d.Controls.Add(prog_id, ctrl_name, True)
                ctrl.Left = left;  ctrl.Top  = top
                ctrl.Width = w;    ctrl.Height = h
                if "Label" in prog_id or "CommandButton" in prog_id or "CheckBox" in prog_id:
                    ctrl.Caption = cap
                if "ListBox" in prog_id and isinstance(extra, dict):
                    ctrl.ColumnCount  = extra.get("cols", 1)
                    ctrl.ColumnWidths = extra.get("colw", "")
                if "Label" in prog_id and isinstance(extra, dict):
                    if extra.get("bold"): ctrl.Font.Bold = True
                    ctrl.Font.Size = extra.get("fsize", 10)
            except: pass
        try: d.Caption = caption
        except: pass
        frm.CodeModule.AddFromString(code)
        print(f"  Form: {name}")
    except Exception as e:
        print(f"  WARNING form {name}: {e}")


def setup_vba(wb):
    try:
        vb = wb.api.VBProject
    except Exception as e:
        print(f"  ERROR: Cannot access VBProject.\n"
              f"  Enable 'Trust access to the VBA project object model'\n"
              f"  in Excel -> File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings\n"
              f"  Error: {e}")
        return

    add_module(vb, "modHelpers",      VBA_HELPERS)
    add_module(vb, "modItems",        VBA_ITEMS)
    add_module(vb, "modTransactions", VBA_TXN)
    add_module(vb, "modInventory",    VBA_INV)
    add_module(vb, "modDashboard",    VBA_DASH)
    add_module(vb, "modReports",      VBA_REPORTS)
    add_module(vb, "modMain",         VBA_MAIN)

    try:
        vb.VBComponents("ThisWorkbook").CodeModule.AddFromString(VBA_THISWORKBOOK)
        print("  ThisWorkbook events set")
    except Exception as e:
        print(f"  WARNING ThisWorkbook: {e}")

    add_form(vb, "frmMain", "ניהול מלאי – מכשור ובקרה", 300, 370, CODE_FRMMAIN, [
        ("Forms.Label.1",         "lblTitle",    "ניהול מלאי מכשור ובקרה",10,8, 270,24,{"fsize":13,"bold":True}),
        ("Forms.CommandButton.1", "cmdAddItem",  "הוספת פריט חדש",        30,45, 230,30,{}),
        ("Forms.CommandButton.1", "cmdStockIn",  "קבלת מלאי (IN)",         30,82, 230,30,{}),
        ("Forms.CommandButton.1", "cmdStockOut", "הוצאת מלאי (OUT)",       30,119,230,30,{}),
        ("Forms.CommandButton.1", "cmdAdjust",   "תיקון מלאי",             30,156,230,30,{}),
        ("Forms.CommandButton.1", "cmdSearch",   "חיפוש פריט",             30,193,230,30,{}),
        ("Forms.CommandButton.1", "cmdRefresh",  "רענן לוח בקרה",          30,230,230,30,{}),
        ("Forms.CommandButton.1", "cmdOrders",   "דוח חוסרים",             30,267,230,30,{}),
        ("Forms.CommandButton.1", "cmdClose",    "סגור",                    30,310,230,28,{}),
    ])

    add_form(vb, "frmStockOut", "הוצאת פריט מהמלאי", 480, 500, CODE_FRMSTOCKOUT, [
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
    ])

    add_form(vb, "frmStockIn", "קבלת מלאי למחסן", 480, 485, CODE_FRMSTOCKIN, [
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
    ])

    add_form(vb, "frmAddItem", "הוספת פריט חדש למאגר", 420, 540, CODE_FRMADDITEM, [
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
    ])

    add_form(vb, "frmAdjust", "תיקון מלאי (Adjustment)", 390, 310, CODE_FRMADJUST, [
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
    ])

    add_form(vb, "frmSearch", "חיפוש פריטים", 640, 478, CODE_FRMSEARCH, [
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
    ])

    print("  All VBA injected")


# ════════════════════════════════════════════════════════════════════════════
#  MAIN
# ════════════════════════════════════════════════════════════════════════════

def main():
    print("=== Building SparePartsInventory_v2.xlsm ===")
    enable_vba_access()

    # Kill any leftover Excel processes first
    os.system("taskkill /F /IM excel.exe >nul 2>&1")
    time.sleep(1)

    app = xw.App(visible=True, add_book=False)
    app.display_alerts = False

    try:
        wb = app.books.add()
        print(f"  Workbook created")

        setup_sheets(wb)
        setup_lists(wb)
        setup_settings(wb)
        setup_data_sheets(wb)
        setup_named_ranges(wb)
        setup_dashboard(wb)
        setup_vba(wb)

        print(f"  Saving -> {OUTPUT_FILE}")
        wb.save(OUTPUT_FILE)
        wb.close()
        print("=== SUCCESS ===")
        print(f"  File: {OUTPUT_FILE}")
        print("  Open the file and click 'ייבוא ממהדורה 1' to import your existing data.")

    except Exception as e:
        import traceback
        print(f"ERROR: {e}")
        traceback.print_exc()
    finally:
        app.quit()


if __name__ == "__main__":
    main()
