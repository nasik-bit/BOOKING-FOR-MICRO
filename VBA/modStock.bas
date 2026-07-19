Attribute VB_Name = "modStock"
Option Explicit

Public Sub GenerateCurrentStock()

    Const MOV_COL_LOCATION As Long = 7
    Const MOV_COL_QTY As Long = 8
    Const OPEN_COL_TYPE As Long = 2
    Const OPEN_COL_LOCATION As Long = 3
    Const OPEN_COL_QTY As Long = 4

    Dim wsMov As Worksheet
    Dim wsOpen As Worksheet
    Dim wsStock As Worksheet

    Dim movLast As Long
    Dim openLast As Long

    Dim movData As Variant
    Dim openData As Variant
    Dim outData(1 To 5, 1 To 2) As Variant

    Dim totals(1 To 4) As Double
    Dim grandTotal As Double

    Dim r As Long
    Dim idx As Long

    Set wsMov = ThisWorkbook.Worksheets(SHEET_MOVEMENT)
    Set wsOpen = ThisWorkbook.Worksheets(SHEET_OPENING)
    Set wsStock = CreateSheet(SHEET_STOCK)

    wsStock.Cells.Clear
    wsStock.Range("A1:B1").Value = Array("Location", "Current Stock")

    movLast = LastRow(wsMov, 1)
    If movLast >= 2 Then
        movData = wsMov.Range("A2:I" & movLast).Value2

        For r = 1 To UBound(movData, 1)
            idx = LocationIndex(movData(r, MOV_COL_LOCATION))
            If idx > 0 Then totals(idx) = totals(idx) + Nz(movData(r, MOV_COL_QTY))
        Next r
    End If

    openLast = LastRow(wsOpen, 1)
    If openLast >= 2 Then
        openData = wsOpen.Range("A2:D" & openLast).Value2

        For r = 1 To UBound(openData, 1)
            idx = LocationIndex(openData(r, OPEN_COL_LOCATION))
            If idx > 0 Then
                totals(idx) = totals(idx) + GetOpeningImpact(openData(r, OPEN_COL_TYPE), Nz(openData(r, OPEN_COL_QTY)))
            End If
        Next r
    End If

    outData(1, 1) = "M&M"
    outData(2, 1) = "NASHIK"
    outData(3, 1) = "PUNE"
    outData(4, 1) = "CK"
    outData(5, 1) = "Grand Total"

    outData(1, 2) = totals(1)
    outData(2, 2) = totals(2)
    outData(3, 2) = totals(3)
    outData(4, 2) = totals(4)

    grandTotal = totals(1) + totals(2) + totals(3) + totals(4)
    outData(5, 2) = grandTotal

    wsStock.Range("A2:B6").Value = outData
    wsStock.Columns("A:B").AutoFit

    MsgBox "Current Stock Generated Successfully.", vbInformation

End Sub

Private Function LocationIndex(ByVal v As Variant) As Long

    Select Case UCase$(Trim$(CStr(v)))
        Case "M&M"
            LocationIndex = 1
        Case "NASHIK"
            LocationIndex = 2
        Case "PUNE"
            LocationIndex = 3
        Case "CK"
            LocationIndex = 4
        Case Else
            LocationIndex = 0
    End Select

End Function

Private Function GetOpeningImpact(ByVal TransactionType As Variant, ByVal Qty As Double) As Double

    Select Case UCase$(Trim$(CStr(TransactionType)))
        Case "OPENING BALANCE", "PURCHASE", "MOVEMENT IN"
            GetOpeningImpact = Abs(Qty)
        Case "MOVEMENT OUT", "SCRAP"
            GetOpeningImpact = -Abs(Qty)
        Case Else
            GetOpeningImpact = Qty
    End Select

End Function
