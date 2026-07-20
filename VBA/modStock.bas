Attribute VB_Name = "modStock"
Option Explicit

Private Const STOCK_MAX_LONG As Long = 2147483647

Public Sub GenerateCurrentStock()

    Const LOC_COUNT As Long = 4
    Const OUT_ROW_COUNT As Long = 5
    Const MOV_COL_DATE As Long = 1
    Const MOV_COL_LOCATION As Long = 7
    Const MOV_COL_QTY As Long = 8
    Const OPEN_COL_DATE As Long = 1
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
    Dim outData(1 To OUT_ROW_COUNT, 1 To 2) As Variant

    Dim totals(1 To LOC_COUNT) As Double
    Dim grandTotal As Double
    Dim earliestOpenDate As Long
    Dim hasOpeningDate As Boolean

    Dim r As Long
    Dim idx As Long
    Dim dKey As Long
    Dim moveDict As Object
    Dim moveKeys() As Variant
    Dim moveCount As Long
    Dim moveArr As Variant
    Dim i As Long

    Set wsMov = ThisWorkbook.Worksheets(SHEET_MOVEMENT)
    Set wsOpen = ThisWorkbook.Worksheets(SHEET_OPENING)
    Set wsStock = CreateSheet(SHEET_STOCK)
    earliestOpenDate = STOCK_MAX_LONG

    wsStock.Cells.Clear
    wsStock.Range("A1:B1").Value = Array("Location", "Current Stock")

    openLast = LastRow(wsOpen, 1)
    If openLast >= 2 Then
        openData = wsOpen.Range("A2:D" & openLast).Value2

        For r = 1 To UBound(openData, 1)
            idx = LocationIndex(openData(r, OPEN_COL_LOCATION))
            If idx > 0 And IsDate(openData(r, OPEN_COL_DATE)) Then
                dKey = CLng(CDate(openData(r, OPEN_COL_DATE)))
                If Not hasOpeningDate Or dKey < earliestOpenDate Then
                    earliestOpenDate = dKey
                    hasOpeningDate = True
                End If
            End If
        Next r

        For r = 1 To UBound(openData, 1)
            idx = LocationIndex(openData(r, OPEN_COL_LOCATION))
            If idx > 0 And hasOpeningDate Then
                If IsDate(openData(r, OPEN_COL_DATE)) Then
                    dKey = CLng(CDate(openData(r, OPEN_COL_DATE)))
                Else
                    dKey = 0
                End If

                If dKey = earliestOpenDate Then
                    totals(idx) = totals(idx) + GetOpeningImpact(openData(r, OPEN_COL_TYPE), Nz(openData(r, OPEN_COL_QTY)))
                End If
            End If
        Next r
    End If

    movLast = LastRow(wsMov, 1)
    If movLast >= 2 Then
        movData = wsMov.Range("A2:H" & movLast).Value2
        Set moveDict = CreateObject("Scripting.Dictionary")

        For r = 1 To UBound(movData, 1)
            If IsDate(movData(r, MOV_COL_DATE)) Then
                idx = LocationIndex(movData(r, MOV_COL_LOCATION))
                If idx > 0 Then
                    dKey = CLng(CDate(movData(r, MOV_COL_DATE)))
                    If (Not hasOpeningDate) Or dKey >= earliestOpenDate Then
                        If Not moveDict.Exists(dKey) Then
                            moveDict(dKey) = StockZeroDeltas(LOC_COUNT)
                        End If
                        moveArr = moveDict(dKey)
                        moveArr(idx - 1) = moveArr(idx - 1) + Nz(movData(r, MOV_COL_QTY))
                        moveDict(dKey) = moveArr
                    End If
                End If
            End If
        Next r

        moveCount = moveDict.Count
        If moveCount > 0 Then
            moveKeys = moveDict.Keys
            StockQuickSort moveKeys, 0, moveCount - 1

            For r = 0 To moveCount - 1
                moveArr = moveDict(moveKeys(r))
                For i = 1 To LOC_COUNT
                    totals(i) = totals(i) + moveArr(i - 1)
                Next i
            Next r
        End If
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

Private Sub StockQuickSort(ByRef arr() As Variant, ByVal lo As Long, ByVal hi As Long)

    Dim pivot As Variant
    Dim i As Long
    Dim j As Long
    Dim tmp As Variant

    If lo >= hi Then Exit Sub

    pivot = arr((lo + hi) \ 2)
    i = lo
    j = hi

    Do While i <= j
        Do While arr(i) < pivot
            i = i + 1
        Loop
        Do While arr(j) > pivot
            j = j - 1
        Loop
        If i <= j Then
            tmp = arr(i)
            arr(i) = arr(j)
            arr(j) = tmp
            i = i + 1
            j = j - 1
        End If
    Loop

    If lo < j Then StockQuickSort arr, lo, j
    If i < hi Then StockQuickSort arr, i, hi

End Sub

Private Function StockZeroDeltas(ByVal count As Long) As Variant

    Dim arr() As Double

    ReDim arr(0 To count - 1)
    StockZeroDeltas = arr

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
