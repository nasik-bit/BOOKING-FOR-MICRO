Attribute VB_Name = "modUtility"
Option Explicit

Public Const SHEET_BOOKING As String = "BOOKING_DATA"
Public Const SHEET_MOVEMENT As String = "MOVEMENT"
Public Const SHEET_DATE_WISE As String = "DATE_WISE_STOCK"
Public Const SHEET_STOCK As String = "CURRENT_STOCK"
Public Const SHEET_VALIDATION As String = "VALIDATION_LOG"

Public Type TAppState
    ScreenUpdating As Boolean
    EnableEvents As Boolean
    DisplayAlerts As Boolean
    Calculation As XlCalculation
    StatusBar As Variant
End Type

Public Sub BeginApp(ByRef state As TAppState, Optional ByVal statusText As String = "Processing...")
    With Application
        state.ScreenUpdating = .ScreenUpdating
        state.EnableEvents = .EnableEvents
        state.DisplayAlerts = .DisplayAlerts
        state.Calculation = .Calculation
        state.StatusBar = .StatusBar

        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
        .StatusBar = statusText
    End With
End Sub

Public Sub EndApp(ByRef state As TAppState)
    With Application
        .ScreenUpdating = state.ScreenUpdating
        .EnableEvents = state.EnableEvents
        .DisplayAlerts = state.DisplayAlerts
        .Calculation = state.Calculation
        .StatusBar = state.StatusBar
    End With
End Sub

Public Function ws(ByVal sheetName As String) As Worksheet
    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Worksheets(sheetName)
    Exit Function
ErrHandler:
    Err.Raise vbObjectError + 1000, "modUtility.ws", "Worksheet not found: " & sheetName
End Function

Public Function TryGetWorksheet(ByVal sheetName As String, ByRef outSheet As Worksheet) As Boolean
    On Error GoTo NotFound
    Set outSheet = ThisWorkbook.Worksheets(sheetName)
    TryGetWorksheet = True
    Exit Function
NotFound:
    Set outSheet = Nothing
    TryGetWorksheet = False
End Function

Public Function CreateSheet(ByVal sheetName As String) As Worksheet
    Dim target As Worksheet
    If TryGetWorksheet(sheetName, target) Then
        Set CreateSheet = target
        Exit Function
    End If

    Set CreateSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    CreateSheet.Name = sheetName
End Function

Public Sub ClearSheet(ByVal target As Worksheet, Optional ByVal keepHeader As Boolean = False)
    If keepHeader Then
        If LastRow(target, 1) > 1 Then
            target.Rows("2:" & target.Rows.Count).ClearContents
        End If
    Else
        target.Cells.Clear
    End If
End Sub

Public Function LastRow(ByVal target As Worksheet, ByVal colIndex As Long) As Long
    Dim r As Long
    r = target.Cells(target.Rows.Count, colIndex).End(xlUp).Row
    If r < 1 Then r = 1
    LastRow = r
End Function

Public Function GetColumn(ByVal target As Worksheet, ByVal headerName As String) As Long
    Dim f As Range
    Set f = target.Rows(1).Find(What:=headerName, _
                                LookIn:=xlValues, _
                                LookAt:=xlWhole, _
                                SearchOrder:=xlByColumns, _
                                SearchDirection:=xlNext, _
                                MatchCase:=False)
    If f Is Nothing Then
        GetColumn = 0
    Else
        GetColumn = f.Column
    End If
End Function

Public Function GetColumnAny(ByVal target As Worksheet, ByVal headerNames As Variant) As Long
    Dim i As Long
    For i = LBound(headerNames) To UBound(headerNames)
        GetColumnAny = GetColumn(target, CStr(headerNames(i)))
        If GetColumnAny > 0 Then Exit Function
    Next i
    GetColumnAny = 0
End Function

Public Function NormalizeText(ByVal v As Variant) As String
    NormalizeText = Trim$(UCase$(CStr(v)))
End Function

Public Function GetPrefix(ByVal cnNo As Variant) As String
    Dim s As String
    s = NormalizeText(cnNo)
    If Len(s) >= 4 Then
        GetPrefix = Left$(s, 4)
    Else
        GetPrefix = s
    End If
End Function

Public Function SafeNumber(ByVal v As Variant, Optional ByVal defaultValue As Double = 0) As Double
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        SafeNumber = defaultValue
        Exit Function
    End If

    If Len(Trim$(CStr(v))) = 0 Then
        SafeNumber = defaultValue
    ElseIf IsNumeric(v) Then
        SafeNumber = CDbl(v)
    Else
        SafeNumber = defaultValue
    End If
End Function

Public Function FDate(ByVal v As Variant) As Variant
    On Error GoTo FailParse

    Dim txt As String
    Dim dt As Date
    Dim y As Long
    Dim m As Long
    Dim d As Long
    Dim serialPart As Long
    Dim parts() As String
    Dim sep As String

    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        FDate = Empty
        Exit Function
    End If

    If VarType(v) = vbDate Then
        dt = DateSerial(Year(v), Month(v), Day(v))
        FDate = dt
        Exit Function
    End If

    If IsNumeric(v) Then
        If CDbl(v) > 0 Then
            serialPart = CLng(Fix(CDbl(v)))
            dt = DateSerial(1899, 12, 30) + serialPart
            FDate = dt
            Exit Function
        End If
        FDate = Empty
        Exit Function
    End If

    txt = Trim$(CStr(v))
    If Len(txt) = 0 Then
        FDate = Empty
        Exit Function
    End If

    txt = Replace$(txt, "T", " ")
    If InStr(1, txt, " ", vbTextCompare) > 0 Then
        txt = Split(txt, " ")(0)
    End If

    If InStr(txt, "-") > 0 Then
        sep = "-"
    ElseIf InStr(txt, "/") > 0 Then
        sep = "/"
    Else
        sep = ""
    End If

    If Len(sep) > 0 Then
        parts = Split(txt, sep)
        If UBound(parts) = 2 Then
            If Len(parts(0)) = 4 Then
                y = CLng(parts(0))
                m = CLng(parts(1))
                d = CLng(parts(2))
            Else
                d = CLng(parts(0))
                m = CLng(parts(1))
                y = CLng(parts(2))
                If y < 100 Then y = y + 2000
            End If

            dt = DateSerial(y, m, d)
            FDate = dt
            Exit Function
        End If
    End If

    If IsDate(txt) Then
        dt = DateValue(txt)
        FDate = dt
        Exit Function
    End If

FailParse:
    FDate = Empty
End Function

Public Function DateKey(ByVal v As Variant) As Long
    Dim dt As Variant
    dt = FDate(v)
    If IsEmpty(dt) Then
        DateKey = 0
    Else
        DateKey = CLng(Fix(CDbl(dt)))
    End If
End Function

Public Function NextDate(ByVal v As Variant) As Variant
    Dim dt As Variant
    dt = FDate(v)
    If IsEmpty(dt) Then
        NextDate = Empty
    Else
        NextDate = DateAdd("d", 1, dt)
    End If
End Function

Public Function KeyToDate(ByVal serialKey As Long) As Date
    KeyToDate = DateSerial(1899, 12, 30) + serialKey
End Function

Public Function NormalizeLocation(ByVal v As Variant) As String
    Dim s As String
    s = NormalizeText(v)
    s = Replace$(s, ".", "")
    s = Replace$(s, "  ", " ")

    Select Case s
        Case "CK"
            NormalizeLocation = "CK"
        Case "PUNE"
            NormalizeLocation = "PUNE"
        Case "NASHIK"
            NormalizeLocation = "NASHIK"
        Case "M&M", "MAHINDRA", "MAHINDRA & MAHINDRA LTD", "MAHINDRA & MAHINDRA LTD."
            NormalizeLocation = "MAHINDRA"
        Case Else
            NormalizeLocation = ""
    End Select
End Function

Public Function IsValidLocationCode(ByVal loc As String) As Boolean
    Select Case loc
        Case "CK", "PUNE", "NASHIK", "MAHINDRA"
            IsValidLocationCode = True
        Case Else
            IsValidLocationCode = False
    End Select
End Function

Public Function IsMahindraConsignee(ByVal v As Variant) As Boolean
    IsMahindraConsignee = (NormalizeText(v) = "MAHINDRA & MAHINDRA LTD.")
End Function

Public Function IsValidPackingType(ByVal v As Variant) As Boolean
    Dim s As String
    s = NormalizeText(v)
    IsValidPackingType = (s = "PLASTIC BIN" Or s = "EMPTY BIN")
End Function

Public Function CanonicalPackingType(ByVal v As Variant) As String
    Dim s As String
    s = NormalizeText(v)
    If s = "PLASTIC BIN" Then
        CanonicalPackingType = "Plastic Bin"
    ElseIf s = "EMPTY BIN" Then
        CanonicalPackingType = "Empty Bin"
    Else
        CanonicalPackingType = ""
    End If
End Function

Public Function IsIgnoredItemCode(ByVal v As Variant) As Boolean
    Select Case NormalizeText(v)
        Case "0304CP500120N", _
             "0309DAV00870N", _
             "03039DAV00880", _
             "0306EAV00320N", _
             "0305GAV00060N", _
             "0309DAV00860N"
            IsIgnoredItemCode = True
        Case Else
            IsIgnoredItemCode = False
    End Select
End Function

Public Function LocationSortIndex(ByVal loc As String) As Long
    Select Case loc
        Case "CK": LocationSortIndex = 1
        Case "PUNE": LocationSortIndex = 2
        Case "NASHIK": LocationSortIndex = 3
        Case "MAHINDRA": LocationSortIndex = 4
        Case Else: LocationSortIndex = 99
    End Select
End Function

Public Sub QuickSortVariant(ByRef arr() As Variant, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long
    Dim j As Long
    Dim pivot As Variant
    Dim tmp As Variant

    If lo >= hi Then Exit Sub
    i = lo
    j = hi
    pivot = arr((lo + hi) \ 2)

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

    If lo < j Then QuickSortVariant arr, lo, j
    If i < hi Then QuickSortVariant arr, i, hi
End Sub

Public Sub ResetValidationLog()
    Dim wsVal As Worksheet
    Set wsVal = CreateSheet(SHEET_VALIDATION)
    wsVal.Cells.Clear
    wsVal.Range("A1:G1").Value = Array("Stage", "Source Row", "CN No", "Invoice Number", "Issue", "Details", "Logged At")
    wsVal.Range("A1:G1").Font.Bold = True
End Sub

Public Sub AppendValidationRows(ByRef logData As Variant, ByVal logCount As Long)
    Dim wsVal As Worksheet
    Dim nextRow As Long
    Dim writeArr() As Variant
    Dim r As Long
    Dim c As Long

    If logCount <= 0 Then Exit Sub

    Set wsVal = CreateSheet(SHEET_VALIDATION)
    If Len(Trim$(CStr(wsVal.Cells(1, 1).Value))) = 0 Then
        ResetValidationLog
    End If

    ReDim writeArr(1 To logCount, 1 To 7)
    For r = 1 To logCount
        For c = 1 To 7
            writeArr(r, c) = logData(r, c)
        Next c
    Next r

    nextRow = LastRow(wsVal, 1) + 1
    wsVal.Range("A" & nextRow).Resize(logCount, 7).Value = writeArr
End Sub
