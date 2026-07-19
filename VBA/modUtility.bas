Attribute VB_Name = "modUtility"
Option Explicit

Public Const SHEET_BOOKING As String = "BOOKING_DATA"
Public Const SHEET_DELIVERY As String = "DELIVERY"
Public Const SHEET_MOVEMENT As String = "MOVEMENT"
Public Const SHEET_MASTER As String = "MASTER"
Public Const SHEET_STOCK As String = "CURRENT_STOCK"
Public Const SHEET_OPENING As String = "OPENING_BALANCE"

Public Sub AppStart()
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
        .StatusBar = "Processing..."
    End With
End Sub

Public Sub AppEnd()
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
        .Calculation = xlCalculationAutomatic
        .StatusBar = False
    End With
End Sub

Public Function LastRow(ws As Worksheet, Col As Variant) As Long
    LastRow = ws.Cells(ws.Rows.Count, Col).End(xlUp).Row
End Function

Public Function GetColumn(ws As Worksheet, HeaderName As String) As Long
    Dim c As Range
    Set c = ws.Rows(1).Find(HeaderName, LookIn:=xlValues, LookAt:=xlWhole)
    If Not c Is Nothing Then
        GetColumn = c.Column
    Else
        GetColumn = 0
    End If
End Function

Public Function FDate(v As Variant) As Variant
    If IsDate(v) Then
        FDate = CDate(v)
    Else
        FDate = Empty
    End If
End Function

Public Sub ClearDataKeepHeader(ws As Worksheet)
    If ws.Rows.Count > 1 Then
        ws.Rows("2:" & ws.Rows.Count).ClearContents
    End If
End Sub

Public Function FindDeliveryRow(CNNo As String) As Long
    Dim ws As Worksheet
    Dim f As Range
    Set ws = ThisWorkbook.Worksheets(SHEET_DELIVERY)

    Set f = ws.Columns(1).Find(What:=CNNo, LookIn:=xlValues, LookAt:=xlWhole)

    If Not f Is Nothing Then
        FindDeliveryRow = f.Row
    Else
        FindDeliveryRow = 0
    End If
End Function

Public Sub WriteLog(ByVal Msg As String)
    Debug.Print Format(Now, "dd-mm-yyyy hh:nn:ss") & " : " & Msg
End Sub
