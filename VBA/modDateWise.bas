Attribute VB_Name = "modDateWise"
Option Explicit

' ===========================================================================
' modDateWise  –  Generate DATE_WISE_STOCK sheet
'
' Output layout (ROWS_PER_DATE rows per date):
'   Date | Location   | Opening | In | Out | Closing
'   -------------------------------------------------
'   date | M&M        | ...
'   date | NASHIK     | ...
'   date | PUNE       | ...
'   date | CK         | ...
'   date | GRAND TOTAL| ...      ← bold + background colour
'   (blank spacer row)
'
' Business Rules:
'   1  Seed opening balance from OPENING_BALANCE sheet.
'   2  Read all MOVEMENT records (columns A:H).
'   3  Sort event dates ascending (QuickSort, O(n log n)).
'   4  Compute Opening / In / Out / Closing per location per date.
'   5  Next day Opening = Previous day Closing.
'   6  Add GRAND TOTAL row after every date.
'   7  Add blank row after every GRAND TOTAL.
'   8  Dictionary ensures no duplicate date accumulations.
'   9  Blank / non-date movement records are silently skipped.
'  10  Supports any number of dates (dynamic arrays + Dictionary).
'
' Performance:
'   Arrays for all bulk reads and writes.
'   Scripting.Dictionary for O(1) per-date lookup.
'   No Select / Activate in the data-processing path.
'   Single range write for all output data.
'   Supports 100,000+ MOVEMENT records.
'
' Validation:
'   Closing < 0  →  highlight cell red; write entry to ERROR_LOG sheet.
'   Macro is never stopped for negative stock.
' ===========================================================================

' --- Error-log sheet name ------------------------------------------------
Private Const SHEET_ERROR_LOG As String = "ERROR_LOG"

' --- MOVEMENT column offsets (resolved at runtime via GetColumn) ---------
Private MOV_DATE As Long   ' Col "Event Date"
Private MOV_LOC  As Long   ' Col "Location"
Private MOV_QTY  As Long   ' Col "Packets" (positive = IN, negative = OUT)

' --- OPENING_BALANCE column offsets (resolved at runtime via GetColumn) --
Private OPEN_DATE As Long  ' Col "Date"
Private OPEN_TYPE As Long  ' Col "Transaction Type"
Private OPEN_LOC  As Long  ' Col "Location"
Private OPEN_QTY  As Long  ' Col "Qty"

' --- Location constants (1-based) ----------------------------------------
Private Const LOC_COUNT As Long = 4  ' M&M=1  NASHIK=2  PUNE=3  CK=4

' --- Output geometry ------------------------------------------------------
Private Const OUT_COLS      As Long = 6  ' Date|Location|Opening|In|Out|Closing
Private Const ROWS_PER_DATE As Long = 6  ' 4 locations + 1 GRAND TOTAL + 1 blank

' --- Cell colours (precomputed RGB values) --------------------------------
'   CLR_HEADER   = RGB(173, 216, 230)  light steel blue
'   CLR_GRANDTOT = RGB(255, 255, 153)  light yellow
'   CLR_NEGATIVE = RGB(255,   0,   0)  red
Private Const CLR_HEADER   As Long = 15128749
Private Const CLR_GRANDTOT As Long = 10092543
Private Const CLR_NEGATIVE As Long = 255
Private Const DATEWISE_MAX_LONG As Long = 2147483647

' ===========================================================================
' GenerateDateWiseStock  –  Public entry point
' ===========================================================================
Public Sub GenerateDateWiseStock()

    ' Worksheet references
    Dim wsOpen As Worksheet
    Dim wsMov  As Worksheet
    Dim wsDW   As Worksheet

    ' Bulk data arrays
    Dim openData  As Variant
    Dim movData   As Variant
    Dim outData() As Variant

    ' Running totals seeded from OPENING_BALANCE; updated after every date
    ' (index 1=M&M, 2=NASHIK, 3=PUNE, 4=CK)
    Dim runTotals(1 To LOC_COUNT) As Double

    ' dateDeltaDict(CLng(dateSerial)) = 8-element Double array:
    '   slots 0,1 = In / Out for M&M
    '         2,3 = In / Out for NASHIK
    '         4,5 = In / Out for PUNE
    '         6,7 = In / Out for CK
    Dim dateDeltaDict As Object
    Dim dateKeys()    As Variant
    Dim dateCount     As Long
    Dim outRow        As Long

    ' Loop / work variables
    Dim r       As Long
    Dim d       As Long
    Dim i       As Long
    Dim idx     As Long
    Dim locOff  As Long
    Dim dKey    As Long
    Dim tempArr As Variant
    Dim qty     As Double
    Dim hasOpeningDate As Boolean
    Dim earliestOpenDate As Long

    ' Per-date accumulators (reused on each date iteration)
    Dim locOpening(1 To LOC_COUNT) As Double
    Dim locIn(1 To LOC_COUNT)      As Double
    Dim locOut(1 To LOC_COUNT)     As Double
    Dim locClosing(1 To LOC_COUNT) As Double

    Dim openLast As Long
    Dim movLast  As Long
    Dim openReadCols As Long
    Dim movReadCols  As Long

    AppStart
    On Error GoTo ErrHandler

    ' ------------------------------------------------------------------
    ' Acquire worksheets
    ' ------------------------------------------------------------------
    Set wsOpen = ThisWorkbook.Worksheets(SHEET_OPENING)
    Set wsMov  = ThisWorkbook.Worksheets(SHEET_MOVEMENT)
    Set wsDW   = CreateSheet(SHEET_DATE_WISE)
    earliestOpenDate = DATEWISE_MAX_LONG

    ' ------------------------------------------------------------------
    ' Resolve column positions dynamically (no hardcoded column numbers)
    ' ------------------------------------------------------------------
    MOV_DATE = GetColumn(wsMov, "Event Date")
    MOV_LOC = GetColumn(wsMov, "Location")
    MOV_QTY = GetColumn(wsMov, "Packets")
    movReadCols = WorksheetFunction.Max(MOV_DATE, MOV_LOC, MOV_QTY)

    OPEN_DATE = GetColumn(wsOpen, "Date")
    OPEN_TYPE = GetColumn(wsOpen, "Transaction Type")
    OPEN_LOC = GetColumn(wsOpen, "Location")
    OPEN_QTY = GetColumn(wsOpen, "Qty")
    openReadCols = WorksheetFunction.Max(OPEN_DATE, OPEN_TYPE, OPEN_LOC, OPEN_QTY)

    If MOV_DATE = 0 Or MOV_LOC = 0 Or MOV_QTY = 0 Then
        Err.Raise vbObjectError + 1, "modDateWise.GenerateDateWiseStock", _
            "MOVEMENT sheet is missing one of the required headers: Event Date, Location, Packets."
    End If

    If OPEN_DATE = 0 Or OPEN_TYPE = 0 Or OPEN_LOC = 0 Or OPEN_QTY = 0 Then
        Err.Raise vbObjectError + 2, "modDateWise.GenerateDateWiseStock", _
            "OPENING_BALANCE sheet is missing one of the required headers: Date, Transaction Type, Location, Qty."
    End If

    wsDW.Cells.Clear   ' Wipe content AND formatting from any prior run

    ' ==================================================================
    ' Step 1  –  Seed running totals from earliest OPENING_BALANCE date only
    ' ==================================================================
    openLast = LastRow(wsOpen, OPEN_DATE)
    If openLast >= 2 Then
        openData = wsOpen.Range(wsOpen.Cells(2, 1), wsOpen.Cells(openLast, openReadCols)).Value2
        For r = 1 To UBound(openData, 1)
            idx = DW_LocationIndex(openData(r, OPEN_LOC))
            If idx > 0 And IsValidDateValue(openData(r, OPEN_DATE)) Then
                dKey = CLng(CDate(openData(r, OPEN_DATE)))
                If Not hasOpeningDate Or dKey < earliestOpenDate Then
                    earliestOpenDate = dKey
                    hasOpeningDate = True
                End If
            End If
        Next r

        For r = 1 To UBound(openData, 1)
            idx = DW_LocationIndex(openData(r, OPEN_LOC))
            If idx > 0 And hasOpeningDate Then
                If IsValidDateValue(openData(r, OPEN_DATE)) Then
                    dKey = CLng(CDate(openData(r, OPEN_DATE)))
                Else
                    dKey = 0
                End If
                If dKey = earliestOpenDate Then
                    runTotals(idx) = runTotals(idx) + _
                        DW_OpeningImpact(openData(r, OPEN_TYPE), Nz(openData(r, OPEN_QTY)))
                End If
            End If
        Next r
    End If

    Set dateDeltaDict = CreateObject("Scripting.Dictionary")
    If hasOpeningDate Then
        dateDeltaDict(earliestOpenDate) = DW_ZeroDateDelta()
    End If

    ' ==================================================================
    ' Step 2  –  Read MOVEMENT into bulk array
    ' ==================================================================
    movLast = LastRow(wsMov, MOV_DATE)
    If movLast >= 2 Then
        movData = wsMov.Range(wsMov.Cells(2, 1), wsMov.Cells(movLast, movReadCols)).Value2

        ' ==================================================================
        ' Step 3  –  Build per-date In / Out delta dictionary (O(n) single pass)
        '
        '   Positive Packets value  →  IN  (slot locOff)
        '   Negative Packets value  →  OUT (slot locOff + 1, stored as Abs)
        ' ==================================================================
        For r = 1 To UBound(movData, 1)
            If IsValidDateValue(movData(r, MOV_DATE)) Then               ' Rule 9: skip blank rows
                idx = DW_LocationIndex(movData(r, MOV_LOC))
                If idx > 0 Then
                    dKey = CLng(CDate(movData(r, MOV_DATE)))
                    If (Not hasOpeningDate) Or dKey >= earliestOpenDate Then
                        If Not dateDeltaDict.Exists(dKey) Then
                            dateDeltaDict(dKey) = DW_ZeroDateDelta()
                        End If
                        qty = Nz(movData(r, MOV_QTY), 0)
                        tempArr = dateDeltaDict(dKey)
                        locOff = (idx - 1) * 2              ' 0-based In/Out slot pair
                        If qty >= 0 Then
                            tempArr(locOff) = tempArr(locOff) + qty       ' IN
                        Else
                            tempArr(locOff + 1) = tempArr(locOff + 1) + Abs(qty)  ' OUT
                        End If
                        dateDeltaDict(dKey) = tempArr         ' Write back (arrays are by value)
                    End If
                End If
            End If
        Next r
    End If

    dateCount = dateDeltaDict.Count
    If dateCount = 0 Then GoTo WriteOutput

    ' ==================================================================
    ' Step 4  –  Sort date keys ascending (QuickSort, O(n log n))
    ' ==================================================================
    dateKeys = dateDeltaDict.Keys
    DW_QuickSort dateKeys, 0, dateCount - 1

    ' ==================================================================
    ' Step 5  –  Accumulate and build output array in memory
    '            Layout: ROWS_PER_DATE rows per date
    '              rows 1-4 : one per location
    '              row  5   : GRAND TOTAL
    '              row  6   : blank spacer
    ' ==================================================================
    ReDim outData(1 To dateCount * ROWS_PER_DATE, 1 To OUT_COLS)

    outRow = 0
    For d = 0 To dateCount - 1
        tempArr = dateDeltaDict(dateKeys(d))

        ' Compute Opening / In / Out / Closing for each location
        For i = 1 To LOC_COUNT
            locOff        = (i - 1) * 2
            locOpening(i) = runTotals(i)              ' Rule 5: Opening = prev Closing
            locIn(i)      = tempArr(locOff)
            locOut(i)     = tempArr(locOff + 1)
            locClosing(i) = locOpening(i) + locIn(i) - locOut(i)
        Next i

        ' 4 location rows
        For i = 1 To LOC_COUNT
            outRow = outRow + 1
            outData(outRow, 1) = CDate(dateKeys(d))
            outData(outRow, 2) = DW_LocationName(i)
            outData(outRow, 3) = locOpening(i)
            outData(outRow, 4) = locIn(i)
            outData(outRow, 5) = locOut(i)
            outData(outRow, 6) = locClosing(i)
        Next i

        ' GRAND TOTAL row (Rule 6)
        outRow = outRow + 1
        outData(outRow, 1) = CDate(dateKeys(d))
        outData(outRow, 2) = "GRAND TOTAL"
        outData(outRow, 3) = locOpening(1) + locOpening(2) + locOpening(3) + locOpening(4)
        outData(outRow, 4) = locIn(1)      + locIn(2)      + locIn(3)      + locIn(4)
        outData(outRow, 5) = locOut(1)     + locOut(2)     + locOut(3)     + locOut(4)
        outData(outRow, 6) = locClosing(1) + locClosing(2) + locClosing(3) + locClosing(4)

        ' Blank spacer row (Rule 7) – outRow incremented; cells remain empty
        outRow = outRow + 1

        ' Roll forward: next day Opening = this day Closing (Rule 5)
        For i = 1 To LOC_COUNT
            runTotals(i) = locClosing(i)
        Next i
    Next d

WriteOutput:
    ' Write header
    wsDW.Range("A1:F1").Value = Array("Date", "Location", "Opening", "In", "Out", "Closing")

    ' Write all data rows in a single range assignment (performance critical)
    If outRow > 0 Then
        wsDW.Range("A2").Resize(outRow, OUT_COLS).Value = outData
    End If

    ' Apply all formatting
    DW_FormatOutput wsDW, dateCount, outRow

    ' Validate and flag negative closing balances
    If outRow > 0 Then
        DW_ValidateAndLog wsDW, outRow
    End If

    AppEnd
    MsgBox "Date-Wise Stock Generated Successfully.", vbInformation
    Exit Sub

ErrHandler:
    AppEnd
    MsgBox "Error in GenerateDateWiseStock: " & Err.Description, vbCritical

End Sub

' ===========================================================================
' DW_FormatOutput  –  Apply all formatting to the DATE_WISE_STOCK sheet
'
' Uses range unions to batch all formatting into a small number of calls.
' Select / Activate is avoided in the data path; FreezePanes is the sole
' exception because Excel's API inherently requires the sheet to be active.
' ===========================================================================
Private Sub DW_FormatOutput(ByVal wsDW As Worksheet, _
                             ByVal dateCount As Long, _
                             ByVal dataRows As Long)

    Dim rngGT    As Range   ' Union of all GRAND TOTAL rows (cols A:F)
    Dim rngData  As Range   ' Union of all non-blank data rows (cols A:F)
    Dim rngBlock As Range   ' One date group's 5 non-blank rows
    Dim rngGTRow As Range   ' One GRAND TOTAL row range
    Dim d        As Long
    Dim baseRow  As Long    ' Sheet row of first location row in a date group
    Dim gtRow    As Long    ' Sheet row of the GRAND TOTAL for a date group

    ' --- Header row (row 1) ---
    With wsDW.Range("A1:F1")
        .Font.Bold           = True
        .Interior.Color      = CLR_HEADER
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle   = xlContinuous
        .Borders.Weight      = xlThin
    End With

    ' --- Date number format on column A ---
    wsDW.Columns("A").NumberFormat = "dd-mm-yyyy"

    If dateCount = 0 Or dataRows = 0 Then GoTo AutoFitAndFreeze

    ' --- Build union ranges across all date groups ---
    For d = 0 To dateCount - 1
        ' baseRow: sheet row of first location row for this date
        baseRow = 2 + d * ROWS_PER_DATE

        ' Non-blank block: 5 rows (4 locations + GRAND TOTAL)
        Set rngBlock = wsDW.Range(wsDW.Cells(baseRow, 1), _
                                  wsDW.Cells(baseRow + 4, OUT_COLS))
        If rngData Is Nothing Then
            Set rngData = rngBlock
        Else
            Set rngData = Union(rngData, rngBlock)
        End If

        ' GRAND TOTAL row: 5th row of the group
        gtRow = baseRow + 4
        Set rngGTRow = wsDW.Range(wsDW.Cells(gtRow, 1), _
                                   wsDW.Cells(gtRow, OUT_COLS))
        If rngGT Is Nothing Then
            Set rngGT = rngGTRow
        Else
            Set rngGT = Union(rngGT, rngGTRow)
        End If
    Next d

    ' --- All borders on non-blank data rows ---
    If Not rngData Is Nothing Then
        rngData.Borders.LineStyle = xlContinuous
        rngData.Borders.Weight    = xlThin
    End If

    ' --- GRAND TOTAL rows: bold + background colour ---
    If Not rngGT Is Nothing Then
        rngGT.Font.Bold      = True
        rngGT.Interior.Color = CLR_GRANDTOT
    End If

AutoFitAndFreeze:
    ' AutoFit columns A:F
    wsDW.Columns("A:F").AutoFit

    ' Freeze top row.
    ' FreezePanes requires the target sheet to be active; ScreenUpdating
    ' is False at this point so the brief activation is invisible to users.
    wsDW.Activate
    wsDW.Range("A2").Select
    ActiveWindow.FreezePanes = False   ' Release any existing freeze first
    ActiveWindow.FreezePanes = True    ' Freeze at row 1 (active cell = A2)

End Sub

' ===========================================================================
' DW_ValidateAndLog  –  Flag Closing < 0 and write to ERROR_LOG
'
' The macro is never stopped for negative stock (Rule 8 / Validation spec).
' Action: highlight the Closing cell red AND append a row to ERROR_LOG.
'
' ERROR_LOG columns: Date | Location | Opening | In | Out | Closing | Error Message
' ===========================================================================
Private Sub DW_ValidateAndLog(ByVal wsDW As Worksheet, ByVal dataRows As Long)

    Dim wsErr       As Worksheet
    Dim sheetData   As Variant
    Dim negData()   As Variant
    Dim negCount    As Long
    Dim n           As Long
    Dim i           As Long
    Dim errStartRow As Long

    ' Read the written output back into an array for fast scanning
    sheetData = wsDW.Range("A2:F" & (1 + dataRows)).Value2

    ' --- First pass: count negatives and highlight Closing cells red ---
    negCount = 0
    For i = 1 To dataRows
        ' Skip blank spacer rows (Location column will be empty)
        If Len(Trim$(CStr(sheetData(i, 2)))) > 0 Then
            If IsNumeric(sheetData(i, 6)) Then
                If CDbl(sheetData(i, 6)) < 0 Then
                    negCount = negCount + 1
                    wsDW.Cells(i + 1, 6).Interior.Color = CLR_NEGATIVE
                End If
            End If
        End If
    Next i

    If negCount = 0 Then Exit Sub

    ' --- Second pass: populate error array for bulk write ---
    ReDim negData(1 To negCount, 1 To 7)
    n = 0
    For i = 1 To dataRows
        If Len(Trim$(CStr(sheetData(i, 2)))) > 0 Then
            If IsNumeric(sheetData(i, 6)) Then
                If CDbl(sheetData(i, 6)) < 0 Then
                    n = n + 1
                    negData(n, 1) = sheetData(i, 1)   ' Date
                    negData(n, 2) = sheetData(i, 2)   ' Location
                    negData(n, 3) = sheetData(i, 3)   ' Opening
                    negData(n, 4) = sheetData(i, 4)   ' In
                    negData(n, 5) = sheetData(i, 5)   ' Out
                    negData(n, 6) = sheetData(i, 6)   ' Closing
                    negData(n, 7) = "Negative Stock"  ' Error Message
                End If
            End If
        End If
    Next i

    ' --- Write to ERROR_LOG in a single range assignment ---
    Set wsErr = CreateSheet(SHEET_ERROR_LOG)

    ' Add header row if the sheet is empty
    If Len(Trim$(CStr(wsErr.Cells(1, 1).Value))) = 0 Then
        wsErr.Range("A1:G1").Value = _
            Array("Date", "Location", "Opening", "In", "Out", "Closing", "Error Message")
        With wsErr.Range("A1:G1")
            .Font.Bold      = True
            .Interior.Color = CLR_HEADER
        End With
    End If

    ' Append errors below any existing rows (supports incremental reruns)
    errStartRow = LastRow(wsErr, 1) + 1
    wsErr.Range("A" & errStartRow).Resize(negCount, 7).Value = negData

    wsErr.Columns("A").NumberFormat = "dd-mm-yyyy"
    wsErr.Columns("A:G").AutoFit

End Sub

' ===========================================================================
' DW_LocationIndex  –  Map location name → 1-based index
' ===========================================================================
Private Function DW_LocationIndex(ByVal v As Variant) As Long
    Select Case UCase$(Trim$(CStr(v)))
        Case "M&M"    : DW_LocationIndex = 1
        Case "NASHIK" : DW_LocationIndex = 2
        Case "PUNE"   : DW_LocationIndex = 3
        Case "CK"     : DW_LocationIndex = 4
        Case Else     : DW_LocationIndex = 0
    End Select
End Function

' ===========================================================================
' DW_LocationName  –  Map 1-based index → location display name
' ===========================================================================
Private Function DW_LocationName(ByVal idx As Long) As String
    Select Case idx
        Case 1  : DW_LocationName = "M&M"
        Case 2  : DW_LocationName = "NASHIK"
        Case 3  : DW_LocationName = "PUNE"
        Case 4  : DW_LocationName = "CK"
        Case Else: DW_LocationName = ""
    End Select
End Function

' ===========================================================================
' DW_OpeningImpact  –  Apply sign rules from OPENING_BALANCE transaction type
'   Opening Balance / Purchase / Movement In  →  +Abs(Qty)
'   Movement Out / Scrap                      →  -Abs(Qty)
'   Physical Adjustment                       →   Qty  (signed as entered)
'   Unknown                                   →   Qty  (pass through signed)
' ===========================================================================
Private Function DW_OpeningImpact(ByVal TransType As Variant, ByVal Qty As Double) As Double
    Select Case UCase$(Trim$(CStr(TransType)))
        Case "OPENING BALANCE", "PURCHASE", "MOVEMENT IN"
            DW_OpeningImpact = Abs(Qty)
        Case "MOVEMENT OUT", "SCRAP"
            DW_OpeningImpact = -Abs(Qty)
        Case "PHYSICAL ADJUSTMENT"
            DW_OpeningImpact = Qty   ' Positive = in, negative = out
        Case Else
            DW_OpeningImpact = Qty   ' Unknown type – pass through signed
    End Select
End Function

Private Function DW_ZeroDateDelta() As Variant

    Dim arr() As Double

    ReDim arr(0 To (LOC_COUNT * 2) - 1)
    DW_ZeroDateDelta = arr

End Function

' ===========================================================================
' DW_QuickSort  –  In-place ascending sort of a 0-based Variant array
' ===========================================================================
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
