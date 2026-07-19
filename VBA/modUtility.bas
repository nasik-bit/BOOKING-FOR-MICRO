Attribute VB_Name = "modUtility"
Option Explicit

Public Const SHEET_BOOKING As String = "BOOKING_DATA"
Public Const SHEET_DELIVERY As String = "DELIVERY"
Public Const SHEET_MOVEMENT As String = "MOVEMENT"
Public Const SHEET_MASTER As String = "MASTER"
Public Const SHEET_STOCK As String = "CURRENT_STOCK"
Public Const SHEET_OPENING As String = "OPENING_BALANCE"

Public Function ws(ByVal SheetName As String) As Worksheet
    Set ws = ThisWorkbook.Worksheets(SheetName)
End Function

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

Public Function CreateSheet(ByVal SheetName As String) As Worksheet
    On Error Resume Next
    Set CreateSheet = ThisWorkbook.Worksheets(SheetName)
    On Error GoTo 0

    If CreateSheet Is Nothing Then
        Set CreateSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        CreateSheet.Name = SheetName
    End If
End Function

Public Sub ClearData(ByVal TargetSheet As Worksheet)
    TargetSheet.Cells.Clear
End Sub

Public Sub ClearDataKeepHeader(ByVal TargetSheet As Worksheet)
    Dim lr As Long

    lr = LastRow(TargetSheet, 1)

    If lr > 1 Then
        TargetSheet.Rows("2:" & lr).ClearContents
    End If
End Sub

Public Function LastRow(ByVal TargetSheet As Worksheet, ByVal Col As Variant) As Long
    LastRow = TargetSheet.Cells(TargetSheet.Rows.Count, Col).End(xlUp).Row
    If LastRow < 1 Then LastRow = 1
End Function

Public Function GetColumn(ByVal TargetSheet As Worksheet, ByVal HeaderName As String) As Long
    Dim c As Range
    Set c = TargetSheet.Rows(1).Find(What:=HeaderName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)

    If Not c Is Nothing Then
        GetColumn = c.Column
    Else
        GetColumn = 0
    End If
End Function

Public Function GetPrefix(ByVal CNNo As Variant) As String
    Dim s As String

    s = Trim$(CStr(CNNo))

    If Len(s) >= 4 Then
        GetPrefix = UCase$(Left$(s, 4))
    Else
        GetPrefix = UCase$(s)
    End If
End Function

Public Function Nz(ByVal v As Variant, Optional ByVal DefaultValue As Double = 0) As Double
    If IsError(v) Or IsNull(v) Or LenB(Trim$(CStr(v))) = 0 Then
        Nz = DefaultValue
    ElseIf IsNumeric(v) Then
        Nz = CDbl(v)
    Else
        Nz = DefaultValue
    End If
End Function

Public Function NextDay(ByVal v As Variant) As Variant
    If IsDate(v) Then
        NextDay = DateAdd("d", 1, CDate(v))
    Else
        NextDay = Empty
    End If
End Function

Public Function FDate(ByVal v As Variant) As Variant
    If IsDate(v) Then
        FDate = CDate(v)
    Else
        FDate = Empty
    End If
End Function

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
