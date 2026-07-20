Attribute VB_Name = "modStock"
Option Explicit

Private Const OUT_COLS As Long = 6

Public Sub GenerateCurrentStock()
    Dim appState As TAppState
    Dim wsDW As Worksheet
    Dim wsStock As Worksheet

    Dim colDate As Long
    Dim colItem As Long
    Dim colLoc As Long
    Dim colClose As Long
    Dim readCols As Long
    Dim lastRw As Long

    Dim data As Variant
    Dim latestDateDict As Object
    Dim latestQtyDict As Object
    Dim itemAggDict As Object

    Dim r As Long
    Dim dateKey As Long
    Dim itemCode As String
    Dim locationCode As String
    Dim closeQty As Double
    Dim keyIL As Variant

    Dim itemKeys() As Variant
    Dim outData() As Variant
    Dim outRow As Long
    Dim i As Long
    Dim rowArr As Variant

    On Error GoTo ErrHandler
    BeginApp appState, "Generating CURRENT_STOCK..."

    Set wsDW = ws(SHEET_DATE_WISE)
    Set wsStock = CreateSheet(SHEET_STOCK)
    wsStock.Cells.Clear
    wsStock.Range("A1:F1").Value = Array("Item Code", "CK", "PUNE", "NASHIK", "MAHINDRA", "Total Packets")
    wsStock.Range("A1:F1").Font.Bold = True

    colDate = GetColumn(wsDW, "Date")
    colItem = GetColumn(wsDW, "Item Code")
    colLoc = GetColumn(wsDW, "Location")
    colClose = GetColumn(wsDW, "Closing Packets")

    If colDate = 0 Or colItem = 0 Or colLoc = 0 Or colClose = 0 Then
        Err.Raise vbObjectError + 1301, "modStock.GenerateCurrentStock", _
                  "DATE_WISE_STOCK requires headers: Date, Item Code, Location, Closing Packets."
    End If

    lastRw = LastRow(wsDW, colDate)
    If lastRw < 2 Then
        EndApp appState
        MsgBox "Current Stock Generated Successfully.", vbInformation
        Exit Sub
    End If

    readCols = WorksheetFunction.Max(colDate, colItem, colLoc, colClose)
    data = wsDW.Range(wsDW.Cells(2, 1), wsDW.Cells(lastRw, readCols)).Value2

    Set latestDateDict = CreateObject("Scripting.Dictionary")
    Set latestQtyDict = CreateObject("Scripting.Dictionary")
    Set itemAggDict = CreateObject("Scripting.Dictionary")

    For r = 1 To UBound(data, 1)
        dateKey = DateKey(data(r, colDate))
        itemCode = NormalizeText(data(r, colItem))
        locationCode = NormalizeLocation(data(r, colLoc))
        closeQty = SafeNumber(data(r, colClose), 0)

        If dateKey = 0 Then GoTo NextRow
        If Len(itemCode) = 0 Then GoTo NextRow
        If Len(locationCode) = 0 Then GoTo NextRow

        keyIL = itemCode & "|" & locationCode
        If Not latestDateDict.Exists(keyIL) Then
            latestDateDict(keyIL) = dateKey
            latestQtyDict(keyIL) = closeQty
        ElseIf dateKey >= CLng(latestDateDict(keyIL)) Then
            latestDateDict(keyIL) = dateKey
            latestQtyDict(keyIL) = closeQty
        End If
NextRow:
    Next r

    For Each keyIL In latestQtyDict.Keys
        itemCode = Split(CStr(keyIL), "|")(0)
        locationCode = Split(CStr(keyIL), "|")(1)
        closeQty = latestQtyDict(keyIL)

        If Not itemAggDict.Exists(itemCode) Then
            itemAggDict(itemCode) = Array(0#, 0#, 0#, 0#) ' CK, PUNE, NASHIK, MAHINDRA
        End If

        rowArr = itemAggDict(itemCode)
        Select Case locationCode
            Case "CK": rowArr(0) = closeQty
            Case "PUNE": rowArr(1) = closeQty
            Case "NASHIK": rowArr(2) = closeQty
            Case "MAHINDRA": rowArr(3) = closeQty
        End Select
        itemAggDict(itemCode) = rowArr
    Next keyIL

    If itemAggDict.Count = 0 Then
        EndApp appState
        MsgBox "Current Stock Generated Successfully.", vbInformation
        Exit Sub
    End If

    itemKeys = itemAggDict.Keys
    QuickSortVariant itemKeys, 0, itemAggDict.Count - 1

    ReDim outData(1 To itemAggDict.Count, 1 To OUT_COLS)
    outRow = 0

    For i = 0 To UBound(itemKeys)
        outRow = outRow + 1
        rowArr = itemAggDict(itemKeys(i))
        outData(outRow, 1) = itemKeys(i)
        outData(outRow, 2) = rowArr(0)
        outData(outRow, 3) = rowArr(1)
        outData(outRow, 4) = rowArr(2)
        outData(outRow, 5) = rowArr(3)
        outData(outRow, 6) = rowArr(0) + rowArr(1) + rowArr(2) + rowArr(3)
    Next i

    wsStock.Range("A2").Resize(outRow, OUT_COLS).Value = outData
    wsStock.Columns("B:F").NumberFormat = "0.00"
    wsStock.Columns("A:F").AutoFit

    EndApp appState
    MsgBox "Current Stock Generated Successfully.", vbInformation
    Exit Sub

ErrHandler:
    EndApp appState
    MsgBox "Error in GenerateCurrentStock: " & Err.Description, vbCritical
End Sub
