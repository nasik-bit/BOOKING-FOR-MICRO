Attribute VB_Name = "modDateWise"
Option Explicit

' ===========================================================================
' modDateWise  –  Generate DATE_WISE_STOCK sheet
'
' Algorithm (array-based, O(n)):
'   1. Read OPENING_BALANCE; compute seed totals per location using sign rules.
'   2. Read MOVEMENT into a Variant array (cols A–H).
'   3. Build a Scripting.Dictionary  dateKey(Long) → delta(0..3)  in one pass.
'   4. QuickSort the unique date keys ascending.
'   5. Walk sorted dates, accumulate running totals, write output in one shot.
'
' Output sheet columns:  Date | M&M | NASHIK | PUNE | CK | Grand Total
' ===========================================================================

Public Sub GenerateDateWiseStock()

    ' ---- column offsets inside OPENING_BALANCE range "B2:D" ----
    Const OPEN_TYPE As Long = 1   ' col B
    Const OPEN_LOC  As Long = 2   ' col C
    Const OPEN_QTY  As Long = 3   ' col D

    ' ---- column indices inside MOVEMENT range "A2:H" ----
    Const MOV_DATE As Long = 1    ' col A  Event Date
    Const MOV_LOC  As Long = 7    ' col G  Location
    Const MOV_QTY  As Long = 8    ' col H  Packets

    ' ---- location indices (1-based) ----
    Const LOC_MM     As Long = 1
    Const LOC_NASHIK As Long = 2
    Const LOC_PUNE   As Long = 3
    Const LOC_CK     As Long = 4

    Dim wsOpen As Worksheet
    Dim wsMov  As Worksheet
    Dim wsDW   As Worksheet

    Dim openData  As Variant
    Dim movData   As Variant
    Dim outData() As Variant

    Dim runTotals(1 To 4) As Double

    Dim dateDeltaDict As Object
    Dim dateKeys()    As Variant
    Dim dateCount     As Long

    Dim r        As Long
    Dim d        As Long
    Dim idx      As Long
    Dim dKey     As Long
    Dim tempArr  As Variant
    Dim openLast As Long
    Dim movLast  As Long
    Dim grand    As Double

    AppStart

    On Error GoTo ErrHandler

    Set wsOpen = ThisWorkbook.Worksheets(SHEET_OPENING)
    Set wsMov  = ThisWorkbook.Worksheets(SHEET_MOVEMENT)
    Set wsDW   = CreateSheet(SHEET_DATE_WISE)

    ClearData wsDW

    ' ---- Header ----
    wsDW.Range("A1:F1").Value = Array("Date", "M&M", "NASHIK", "PUNE", "CK", "Grand Total")

    ' =========================================================
    ' Step 1 – Seed running totals from OPENING_BALANCE
    ' =========================================================
    openLast = LastRow(wsOpen, 1)
    If openLast >= 2 Then
        openData = wsOpen.Range("B2:D" & openLast).Value2
        For r = 1 To UBound(openData, 1)
            idx = DW_LocationIndex(openData(r, OPEN_LOC))
            If idx > 0 Then
                runTotals(idx) = runTotals(idx) + _
                    DW_OpeningImpact(openData(r, OPEN_TYPE), Nz(openData(r, OPEN_QTY)))
            End If
        Next r
    End If

    ' =========================================================
    ' Step 2 – Read MOVEMENT data
    ' =========================================================
    movLast = LastRow(wsMov, 1)
    If movLast < 2 Then GoTo WriteOutput

    movData = wsMov.Range("A2:H" & movLast).Value2

    ' =========================================================
    ' Step 3 – Build per-date delta dictionary  (one pass, O(n))
    ' =========================================================
    Set dateDeltaDict = CreateObject("Scripting.Dictionary")

    For r = 1 To UBound(movData, 1)
        If IsDate(movData(r, MOV_DATE)) Then
            dKey = CLng(CDate(movData(r, MOV_DATE)))
            If Not dateDeltaDict.Exists(dKey) Then
                dateDeltaDict(dKey) = Array(0#, 0#, 0#, 0#)
            End If
            idx = DW_LocationIndex(movData(r, MOV_LOC))
            If idx > 0 Then
                tempArr = dateDeltaDict(dKey)
                tempArr(idx - 1) = tempArr(idx - 1) + Nz(movData(r, MOV_QTY))
                dateDeltaDict(dKey) = tempArr
            End If
        End If
    Next r

    dateCount = dateDeltaDict.Count
    If dateCount = 0 Then GoTo WriteOutput

    ' =========================================================
    ' Step 4 – Sort date keys ascending
    '          QuickSort O(n log n) – practical date range is
    '          unlikely to exceed a few hundred unique values.
    ' =========================================================
    dateKeys = dateDeltaDict.Keys
    DW_QuickSort dateKeys, 0, dateCount - 1

    ' =========================================================
    ' Step 5 – Accumulate and build output array
    ' =========================================================
    ReDim outData(1 To dateCount, 1 To 6)

    For d = 0 To dateCount - 1
        tempArr = dateDeltaDict(dateKeys(d))

        runTotals(LOC_MM)     = runTotals(LOC_MM)     + tempArr(0)
        runTotals(LOC_NASHIK) = runTotals(LOC_NASHIK) + tempArr(1)
        runTotals(LOC_PUNE)   = runTotals(LOC_PUNE)   + tempArr(2)
        runTotals(LOC_CK)     = runTotals(LOC_CK)     + tempArr(3)

        grand = runTotals(LOC_MM) + runTotals(LOC_NASHIK) + _
                runTotals(LOC_PUNE) + runTotals(LOC_CK)

        outData(d + 1, 1) = CDate(dateKeys(d))
        outData(d + 1, 2) = runTotals(LOC_MM)
        outData(d + 1, 3) = runTotals(LOC_NASHIK)
        outData(d + 1, 4) = runTotals(LOC_PUNE)
        outData(d + 1, 5) = runTotals(LOC_CK)
        outData(d + 1, 6) = grand
    Next d

    ' Single range write for performance
    wsDW.Range("A2").Resize(dateCount, 6).Value = outData

WriteOutput:
    wsDW.Columns("A").NumberFormat = "dd-mm-yyyy"
    wsDW.Columns("A:F").AutoFit

    AppEnd
    MsgBox "Date-Wise Stock Generated Successfully.", vbInformation
    Exit Sub

ErrHandler:
    AppEnd
    MsgBox "Error in GenerateDateWiseStock: " & Err.Description, vbCritical

End Sub

' ---------------------------------------------------------------------------
' DW_LocationIndex  –  map location name to 1-based index
' ---------------------------------------------------------------------------
Private Function DW_LocationIndex(ByVal v As Variant) As Long
    Select Case UCase$(Trim$(CStr(v)))
        Case "M&M"    : DW_LocationIndex = 1
        Case "NASHIK" : DW_LocationIndex = 2
        Case "PUNE"   : DW_LocationIndex = 3
        Case "CK"     : DW_LocationIndex = 4
        Case Else     : DW_LocationIndex = 0
    End Select
End Function

' ---------------------------------------------------------------------------
' DW_OpeningImpact  –  apply sign rules from OPENING_BALANCE transaction type
'   Opening Balance / Purchase / Movement In  → +Abs(Qty)
'   Movement Out / Scrap                      → -Abs(Qty)
'   Physical Adjustment                       →  Qty  (signed as entered)
' ---------------------------------------------------------------------------
Private Function DW_OpeningImpact(ByVal TransType As Variant, ByVal Qty As Double) As Double
    Select Case UCase$(Trim$(CStr(TransType)))
        Case "OPENING BALANCE", "PURCHASE", "MOVEMENT IN"
            DW_OpeningImpact = Abs(Qty)
        Case "MOVEMENT OUT", "SCRAP"
            DW_OpeningImpact = -Abs(Qty)
        Case "PHYSICAL ADJUSTMENT"
            DW_OpeningImpact = Qty   ' Signed as entered; positive = in, negative = out
        Case Else
            DW_OpeningImpact = Qty   ' Unknown type – pass through signed
    End Select
End Function

' ---------------------------------------------------------------------------
' DW_QuickSort  –  in-place ascending sort of a 0-based Variant array
' ---------------------------------------------------------------------------
Private Sub DW_QuickSort(ByRef arr() As Variant, ByVal lo As Long, ByVal hi As Long)
    Dim pivot   As Variant
    Dim i       As Long
    Dim j       As Long
    Dim swapVal As Variant

    If lo >= hi Then Exit Sub

    pivot = arr((lo + hi) \ 2)
    i = lo
    j = hi

    Do While i <= j
        Do While arr(i) < pivot : i = i + 1 : Loop
        Do While arr(j) > pivot : j = j - 1 : Loop
        If i <= j Then
            swapVal = arr(i)
            arr(i)  = arr(j)
            arr(j)  = swapVal
            i = i + 1
            j = j - 1
        End If
    Loop

    If lo < j Then DW_QuickSort arr, lo, j
    If i < hi Then DW_QuickSort arr, i, hi
End Sub
