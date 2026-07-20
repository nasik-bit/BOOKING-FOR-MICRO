Attribute VB_Name = "modDateWise"
Option Explicit

Private Const OUT_COLS As Long = 6
Private Const LOG_COLS As Long = 7

Public Sub GenerateDateWiseStock()
    Dim appState As TAppState
    Dim wsMov As Worksheet
    Dim wsDW As Worksheet

    Dim colDate As Long
    Dim colItem As Long
    Dim colLoc As Long
    Dim colPackets As Long
    Dim readCols As Long
    Dim lastRw As Long

    Dim data As Variant
    Dim movDict As Object
    Dim itemLocDict As Object
    Dim runBalDict As Object

    Dim itemLocKeys() As Variant
    Dim itemLocCount As Long

    Dim minDate As Long
    Dim maxDate As Long
    Dim dayCount As Long
    Dim totalRows As Long

    Dim outData() As Variant
    Dim outRow As Long
    Dim logData() As Variant
    Dim logCount As Long

    Dim r As Long
    Dim d As Long
    Dim i As Long

    Dim keyIL As String
    Dim keyDIL As String
    Dim itemCode As String
    Dim locationCode As String
    Dim movement As Double
    Dim openingQty As Double
    Dim closingQty As Double
    Dim dtKey As Long

    On Error GoTo ErrHandler
    BeginApp appState, "Generating DATE_WISE_STOCK..."

    Set wsMov = ws(SHEET_MOVEMENT)
    Set wsDW = CreateSheet(SHEET_DATE_WISE)

    wsDW.Cells.Clear
    wsDW.Range("A1:F1").Value = Array("Date", "Item Code", "Location", "Opening Packets", "Movement Packets", "Closing Packets")
    wsDW.Range("A1:F1").Font.Bold = True

    colDate = GetColumn(wsMov, "Event Date")
    colItem = GetColumn(wsMov, "Item Code")
    colLoc = GetColumn(wsMov, "Location")
    colPackets = GetColumn(wsMov, "Packets")
    If colDate = 0 Or colItem = 0 Or colLoc = 0 Or colPackets = 0 Then
        Err.Raise vbObjectError + 1201, "modDateWise.GenerateDateWiseStock", _
                  "MOVEMENT requires headers: Event Date, Item Code, Location, Packets."
    End If

    lastRw = LastRow(wsMov, colDate)
    If lastRw < 2 Then
        EndApp appState
        MsgBox "Date-Wise Stock Generated Successfully.", vbInformation
        Exit Sub
    End If

    readCols = WorksheetFunction.Max(colDate, colItem, colLoc, colPackets)
    data = wsMov.Range(wsMov.Cells(2, 1), wsMov.Cells(lastRw, readCols)).Value2

    Set movDict = CreateObject("Scripting.Dictionary")
    Set itemLocDict = CreateObject("Scripting.Dictionary")
    Set runBalDict = CreateObject("Scripting.Dictionary")

    ReDim logData(1 To (UBound(data, 1) * 20) + 100, 1 To LOG_COLS)
    logCount = 0

    minDate = 0
    maxDate = 0

    For r = 1 To UBound(data, 1)
        dtKey = DateKey(data(r, colDate))
        itemCode = NormalizeText(data(r, colItem))
        locationCode = NormalizeLocation(data(r, colLoc))
        movement = SafeNumber(data(r, colPackets), 0)

        If dtKey = 0 Then
            AddDWValidation logData, logCount, r + 1, "", "", "Blank Dates", "Event Date is blank/invalid."
            GoTo NextRow
        End If
        If Len(itemCode) = 0 Then
            AddDWValidation logData, logCount, r + 1, "", "", "Invalid Movement", "Item Code is blank."
            GoTo NextRow
        End If
        If Len(locationCode) = 0 Then
            AddDWValidation logData, logCount, r + 1, "", "", "Invalid Movement", "Location is blank/invalid."
            GoTo NextRow
        End If

        keyIL = itemCode & "|" & locationCode
        keyDIL = CStr(dtKey) & "|" & keyIL

        If Not itemLocDict.Exists(keyIL) Then itemLocDict(keyIL) = True
        If movDict.Exists(keyDIL) Then
            movDict(keyDIL) = movDict(keyDIL) + movement
        Else
            movDict(keyDIL) = movement
        End If

        If minDate = 0 Or dtKey < minDate Then minDate = dtKey
        If maxDate = 0 Or dtKey > maxDate Then maxDate = dtKey
NextRow:
    Next r

    itemLocCount = itemLocDict.Count
    If itemLocCount = 0 Or minDate = 0 Then
        If logCount > 0 Then AppendValidationRows logData, logCount
        EndApp appState
        MsgBox "Date-Wise Stock Generated Successfully.", vbInformation
        Exit Sub
    End If

    itemLocKeys = itemLocDict.Keys
    QuickSortVariant itemLocKeys, 0, itemLocCount - 1

    dayCount = (maxDate - minDate) + 1
    totalRows = dayCount * itemLocCount
    ReDim outData(1 To totalRows, 1 To OUT_COLS)

    For i = 0 To itemLocCount - 1
        runBalDict(itemLocKeys(i)) = 0#
    Next i

    outRow = 0
    For d = minDate To maxDate
        For i = 0 To itemLocCount - 1
            keyIL = CStr(itemLocKeys(i))
            keyDIL = CStr(d) & "|" & keyIL

            openingQty = runBalDict(keyIL)
            If movDict.Exists(keyDIL) Then
                movement = movDict(keyDIL)
            Else
                movement = 0#
            End If

            closingQty = openingQty + movement
            runBalDict(keyIL) = closingQty

            outRow = outRow + 1
            outData(outRow, 1) = KeyToDate(d)
            outData(outRow, 2) = Split(keyIL, "|")(0)
            outData(outRow, 3) = Split(keyIL, "|")(1)
            outData(outRow, 4) = openingQty
            outData(outRow, 5) = movement
            outData(outRow, 6) = closingQty

            If closingQty < 0 Then
                AddDWValidation logData, logCount, 0, "", "", "Negative Closing Stock", _
                                "Date=" & Format$(KeyToDate(d), "dd-mm-yyyy") & ", Item=" & Split(keyIL, "|")(0) & ", Location=" & Split(keyIL, "|")(1)
            End If
        Next i
    Next d

    wsDW.Range("A2").Resize(outRow, OUT_COLS).Value = outData
    wsDW.Columns("A:A").NumberFormat = "dd-mm-yyyy"
    wsDW.Columns("D:F").NumberFormat = "0.00"
    wsDW.Columns("A:F").AutoFit

    If logCount > 0 Then
        AppendValidationRows logData, logCount
    End If

    EndApp appState
    MsgBox "Date-Wise Stock Generated Successfully.", vbInformation
    Exit Sub

ErrHandler:
    EndApp appState
    MsgBox "Error in GenerateDateWiseStock: " & Err.Description, vbCritical
End Sub

Private Sub AddDWValidation(ByRef logData As Variant, ByRef logCount As Long, _
                            ByVal sourceRow As Long, ByVal cnNo As String, ByVal invoiceNo As String, _
                            ByVal issueName As String, ByVal details As String)
    If logCount >= UBound(logData, 1) Then
        Err.Raise vbObjectError + 1202, "modDateWise.AddDWValidation", "Validation buffer overflow. Increase LOG allocation."
    End If

    logCount = logCount + 1
    logData(logCount, 1) = "DATE_WISE"
    logData(logCount, 2) = sourceRow
    logData(logCount, 3) = cnNo
    logData(logCount, 4) = invoiceNo
    logData(logCount, 5) = issueName
    logData(logCount, 6) = details
    logData(logCount, 7) = Now
End Sub
