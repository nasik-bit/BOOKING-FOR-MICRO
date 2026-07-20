Attribute VB_Name = "modMovement"
Option Explicit

Private Const MOV_COL_COUNT As Long = 15
Private Const LOG_COL_COUNT As Long = 7

Public Sub GenerateMovement()
    Dim appState As TAppState
    Dim wsBook As Worksheet
    Dim wsMov As Worksheet

    Dim colCN As Long
    Dim colInv As Long
    Dim colBookDate As Long
    Dim colTNDate As Long
    Dim colDNDate As Long
    Dim colPackets As Long
    Dim colPacking As Long
    Dim colItem As Long
    Dim colSource As Long
    Dim colDestination As Long
    Dim colConsignor As Long
    Dim colConsignee As Long

    Dim readCols As Long
    Dim lastRw As Long
    Dim data As Variant

    Dim outData() As Variant
    Dim outCount As Long

    Dim logData() As Variant
    Dim logCount As Long

    Dim invDict As Object
    Dim r As Long

    On Error GoTo ErrHandler
    BeginApp appState, "Generating MOVEMENT..."

    Set wsBook = ws(SHEET_BOOKING)
    Set wsMov = CreateSheet(SHEET_MOVEMENT)

    ResetValidationLog
    InitializeMovementHeader wsMov

    colCN = GetColumn(wsBook, "Consignment Number")
    colInv = GetColumnAny(wsBook, Array("Invoice Number", "Invoice No", "Invoice", "Inv No"))
    colBookDate = GetColumn(wsBook, "Consignment Date")
    colTNDate = GetColumn(wsBook, "Transport Note Date")
    colDNDate = GetColumn(wsBook, "DN Date")
    colPackets = GetColumn(wsBook, "Packets")
    colPacking = GetColumn(wsBook, "Packing Type")
    colItem = GetColumn(wsBook, "Item Code")
    colSource = GetColumn(wsBook, "Source")
    colDestination = GetColumn(wsBook, "Destination")
    colConsignor = GetColumn(wsBook, "Consignor")
    colConsignee = GetColumn(wsBook, "Consignee")

    ValidateBookingHeaders colCN, colInv, colBookDate, colTNDate, colDNDate, _
                           colPackets, colPacking, colItem, colSource, colDestination, colConsignor, colConsignee

    readCols = WorksheetFunction.Max(colCN, colInv, colBookDate, colTNDate, colDNDate, colPackets, _
                                     colPacking, colItem, colSource, colDestination, colConsignor, colConsignee)

    lastRw = LastRow(wsBook, colCN)
    If lastRw < 2 Then
        EndApp appState
        MsgBox "Movement Generated Successfully.", vbInformation
        Exit Sub
    End If

    data = wsBook.Range(wsBook.Cells(2, 1), wsBook.Cells(lastRw, readCols)).Value2

    ReDim outData(1 To (UBound(data, 1) * 6) + 10, 1 To MOV_COL_COUNT)
    ReDim logData(1 To (UBound(data, 1) * 12) + 10, 1 To LOG_COL_COUNT)

    Set invDict = CreateObject("Scripting.Dictionary")

    For r = 1 To UBound(data, 1)
        ProcessBookingRecord data, r, _
                             colCN, colInv, colBookDate, colTNDate, colDNDate, _
                             colPackets, colPacking, colItem, colSource, colDestination, colConsignor, colConsignee, _
                             invDict, _
                             outData, outCount, _
                             logData, logCount
    Next r

    If outCount > 0 Then
        wsMov.Range("A2").Resize(outCount, MOV_COL_COUNT).Value = FirstRows(outData, outCount, MOV_COL_COUNT)
        wsMov.Columns("A:A").NumberFormat = "dd-mm-yyyy"
        wsMov.Columns("H:H").NumberFormat = "0.00"
        wsMov.Columns("A:O").AutoFit
    End If

    If logCount > 0 Then
        AppendValidationRows logData, logCount
    End If

    EndApp appState
    MsgBox "Movement Generated Successfully.", vbInformation
    Exit Sub

ErrHandler:
    EndApp appState
    MsgBox "Error in GenerateMovement: " & Err.Description, vbCritical
End Sub

Private Function FirstRows(ByRef src As Variant, ByVal rowCount As Long, ByVal colCount As Long) As Variant
    Dim dst() As Variant
    Dim r As Long
    Dim c As Long

    ReDim dst(1 To rowCount, 1 To colCount)
    For r = 1 To rowCount
        For c = 1 To colCount
            dst(r, c) = src(r, c)
        Next c
    Next r

    FirstRows = dst
End Function

Private Sub InitializeMovementHeader(ByVal wsMov As Worksheet)
    wsMov.Cells.Clear
    wsMov.Range("A1:O1").Value = Array( _
        "Event Date", _
        "CN No", _
        "Invoice Number", _
        "Prefix", _
        "Event", _
        "Source", _
        "Destination", _
        "Location", _
        "Packets", _
        "Packing Type", _
        "Item Code", _
        "Consignor", _
        "Consignee", _
        "Reference", _
        "Remarks")
    wsMov.Range("A1:O1").Font.Bold = True
End Sub

Private Sub ValidateBookingHeaders(ByVal colCN As Long, ByVal colInv As Long, ByVal colBookDate As Long, _
                                   ByVal colTNDate As Long, ByVal colDNDate As Long, ByVal colPackets As Long, _
                                   ByVal colPacking As Long, ByVal colItem As Long, ByVal colSource As Long, _
                                   ByVal colDestination As Long, ByVal colConsignor As Long, ByVal colConsignee As Long)
    If colCN = 0 Then Err.Raise vbObjectError + 1101, "modMovement.ValidateBookingHeaders", "Missing header: Consignment Number"
    If colInv = 0 Then Err.Raise vbObjectError + 1102, "modMovement.ValidateBookingHeaders", "Missing header: Invoice Number"
    If colBookDate = 0 Then Err.Raise vbObjectError + 1103, "modMovement.ValidateBookingHeaders", "Missing header: Consignment Date"
    If colTNDate = 0 Then Err.Raise vbObjectError + 1104, "modMovement.ValidateBookingHeaders", "Missing header: Transport Note Date"
    If colDNDate = 0 Then Err.Raise vbObjectError + 1105, "modMovement.ValidateBookingHeaders", "Missing header: DN Date"
    If colPackets = 0 Then Err.Raise vbObjectError + 1106, "modMovement.ValidateBookingHeaders", "Missing header: Packets"
    If colPacking = 0 Then Err.Raise vbObjectError + 1107, "modMovement.ValidateBookingHeaders", "Missing header: Packing Type"
    If colItem = 0 Then Err.Raise vbObjectError + 1108, "modMovement.ValidateBookingHeaders", "Missing header: Item Code"
    If colSource = 0 Then Err.Raise vbObjectError + 1109, "modMovement.ValidateBookingHeaders", "Missing header: Source"
    If colDestination = 0 Then Err.Raise vbObjectError + 1110, "modMovement.ValidateBookingHeaders", "Missing header: Destination"
    If colConsignor = 0 Then Err.Raise vbObjectError + 1111, "modMovement.ValidateBookingHeaders", "Missing header: Consignor"
    If colConsignee = 0 Then Err.Raise vbObjectError + 1112, "modMovement.ValidateBookingHeaders", "Missing header: Consignee"
End Sub

Private Sub ProcessBookingRecord(ByRef data As Variant, ByVal r As Long, _
                                 ByVal colCN As Long, ByVal colInv As Long, ByVal colBookDate As Long, _
                                 ByVal colTNDate As Long, ByVal colDNDate As Long, ByVal colPackets As Long, _
                                 ByVal colPacking As Long, ByVal colItem As Long, ByVal colSource As Long, _
                                 ByVal colDestination As Long, ByVal colConsignor As Long, ByVal colConsignee As Long, _
                                 ByRef invDict As Object, _
                                 ByRef outData As Variant, ByRef outCount As Long, _
                                 ByRef logData As Variant, ByRef logCount As Long)
    Dim cn As String
    Dim invoiceNo As String
    Dim prefix As String
    Dim itemCode As String
    Dim packingType As String
    Dim sourceLoc As String
    Dim destLoc As String
    Dim consignor As String
    Dim consignee As String
    Dim packets As Double

    Dim bookingDate As Variant
    Dim tnDate As Variant
    Dim dnDate As Variant

    cn = Trim$(CStr(data(r, colCN)))
    invoiceNo = Trim$(CStr(data(r, colInv)))
    prefix = GetPrefix(cn)
    itemCode = NormalizeText(data(r, colItem))
    packingType = CanonicalPackingType(data(r, colPacking))
    sourceLoc = NormalizeLocation(data(r, colSource))
    destLoc = NormalizeLocation(data(r, colDestination))
    consignor = Trim$(CStr(data(r, colConsignor)))
    consignee = Trim$(CStr(data(r, colConsignee)))
    packets = SafeNumber(data(r, colPackets), 0)

    bookingDate = FDate(data(r, colBookDate))
    tnDate = FDate(data(r, colTNDate))
    dnDate = FDate(data(r, colDNDate))

    If Len(cn) = 0 Then
        AddValidation logData, logCount, "MOVEMENT", r + 1, cn, invoiceNo, "Invalid Movement", "Blank Consignment Number."
        Exit Sub
    End If

    If IsIgnoredItemCode(itemCode) Then Exit Sub

    If Not IsValidPackingType(data(r, colPacking)) Then
        AddValidation logData, logCount, "MOVEMENT", r + 1, cn, invoiceNo, "Invalid Packing Type", _
                      "Only Plastic Bin and Empty Bin are allowed."
        Exit Sub
    End If

    If packets <= 0 Then
        AddValidation logData, logCount, "MOVEMENT", r + 1, cn, invoiceNo, "Invalid Movement", _
                      "Packets must be greater than zero."
        Exit Sub
    End If

    If Len(sourceLoc) = 0 Then
        AddValidation logData, logCount, "MOVEMENT", r + 1, cn, invoiceNo, "Missing Source", "Source location is blank or invalid."
        Exit Sub
    End If

    If Len(destLoc) = 0 Then
        AddValidation logData, logCount, "MOVEMENT", r + 1, cn, invoiceNo, "Missing Destination", "Destination location is blank or invalid."
        Exit Sub
    End If

    If Len(invoiceNo) > 0 Then
        If invDict.Exists(invoiceNo) Then
            AddValidation logData, logCount, "MOVEMENT", r + 1, cn, invoiceNo, "Duplicate Invoice", "Invoice appears more than once."
        Else
            invDict(invoiceNo) = True
        End If
    End If

    If IsEmpty(bookingDate) Then
        AddValidation logData, logCount, "MOVEMENT", r + 1, cn, invoiceNo, "Blank Dates", "Booking date is blank/invalid."
        Exit Sub
    End If

    Select Case prefix
        Case "CPUN"
            EmitMovement outData, outCount, bookingDate, cn, invoiceNo, prefix, "BOOKING", sourceLoc, destLoc, "CK", -packets, packingType, itemCode, consignor, consignee, "Booking", ""
            EmitMovement outData, outCount, bookingDate, cn, invoiceNo, prefix, "BOOKING", sourceLoc, destLoc, "PUNE", packets, packingType, itemCode, consignor, consignee, "Booking", ""

            If IsEmpty(tnDate) Then
                AddValidation logData, logCount, "MOVEMENT", r + 1, cn, invoiceNo, "Blank Dates", "Transport Note Date is blank/invalid."
            Else
                EmitMovement outData, outCount, tnDate, cn, invoiceNo, prefix, "TRANSPORT NOTE", sourceLoc, destLoc, "PUNE", -packets, packingType, itemCode, consignor, consignee, "Transport Note Date", ""
                EmitMovement outData, outCount, NextDate(tnDate), cn, invoiceNo, prefix, "TRANSPORT NOTE +1", sourceLoc, destLoc, "NASHIK", packets, packingType, itemCode, consignor, consignee, "Transport Note Date +1", ""
            End If

            If IsMahindraConsignee(consignee) Then
                If IsEmpty(dnDate) Then
                    AddValidation logData, logCount, "MOVEMENT", r + 1, cn, invoiceNo, "Blank Dates", "Delivery Note Date is blank/invalid for Mahindra consignee."
                Else
                    EmitMovement outData, outCount, dnDate, cn, invoiceNo, prefix, "DELIVERY NOTE", sourceLoc, destLoc, "NASHIK", -packets, packingType, itemCode, consignor, consignee, "Delivery Note Date", ""
                    EmitMovement outData, outCount, dnDate, cn, invoiceNo, prefix, "DELIVERY NOTE", sourceLoc, destLoc, "MAHINDRA", packets, packingType, itemCode, consignor, consignee, "Delivery Note Date", ""
                End If
            End If

        Case "CNSK"
            EmitMovement outData, outCount, bookingDate, cn, invoiceNo, prefix, "BOOKING", sourceLoc, destLoc, "MAHINDRA", -packets, packingType, itemCode, consignor, consignee, "Booking", ""
            EmitMovement outData, outCount, bookingDate, cn, invoiceNo, prefix, "BOOKING", sourceLoc, destLoc, "NASHIK", packets, packingType, itemCode, consignor, consignee, "Booking", ""

            If IsEmpty(tnDate) Then
                AddValidation logData, logCount, "MOVEMENT", r + 1, cn, invoiceNo, "Blank Dates", "Transport Note Date is blank/invalid."
            Else
                EmitMovement outData, outCount, tnDate, cn, invoiceNo, prefix, "TRANSPORT NOTE", sourceLoc, destLoc, "NASHIK", -packets, packingType, itemCode, consignor, consignee, "Transport Note Date", ""
                EmitMovement outData, outCount, NextDate(tnDate), cn, invoiceNo, prefix, "TRANSPORT NOTE +1", sourceLoc, destLoc, "PUNE", packets, packingType, itemCode, consignor, consignee, "Transport Note Date +1", ""
            End If

            If IsEmpty(dnDate) Then
                AddValidation logData, logCount, "MOVEMENT", r + 1, cn, invoiceNo, "Blank Dates", "Delivery Note Date is blank/invalid."
            Else
                EmitMovement outData, outCount, dnDate, cn, invoiceNo, prefix, "DELIVERY NOTE", sourceLoc, destLoc, "PUNE", -packets, packingType, itemCode, consignor, consignee, "Delivery Note Date", ""
                EmitMovement outData, outCount, dnDate, cn, invoiceNo, prefix, "DELIVERY NOTE", sourceLoc, destLoc, "CK", packets, packingType, itemCode, consignor, consignee, "Delivery Note Date", ""
            End If

        Case Else
            AddValidation logData, logCount, "MOVEMENT", r + 1, cn, invoiceNo, "Invalid Movement", "Unsupported CN prefix. Allowed: CPUN, CNSK."
    End Select
End Sub

Private Sub EmitMovement(ByRef outData As Variant, ByRef outCount As Long, _
                         ByVal eventDate As Variant, ByVal cnNo As String, ByVal invoiceNo As String, _
                         ByVal prefix As String, ByVal eventName As String, ByVal sourceLoc As String, _
                         ByVal destLoc As String, ByVal location As String, ByVal packets As Double, _
                         ByVal packingType As String, ByVal itemCode As String, ByVal consignor As String, _
                         ByVal consignee As String, ByVal reference As String, ByVal remarks As String)
    outCount = outCount + 1
    outData(outCount, 1) = eventDate
    outData(outCount, 2) = cnNo
    outData(outCount, 3) = invoiceNo
    outData(outCount, 4) = prefix
    outData(outCount, 5) = eventName
    outData(outCount, 6) = sourceLoc
    outData(outCount, 7) = destLoc
    outData(outCount, 8) = location
    outData(outCount, 9) = packets
    outData(outCount, 10) = packingType
    outData(outCount, 11) = itemCode
    outData(outCount, 12) = consignor
    outData(outCount, 13) = consignee
    outData(outCount, 14) = reference
    outData(outCount, 15) = remarks
End Sub

Private Sub AddValidation(ByRef logData As Variant, ByRef logCount As Long, _
                          ByVal stageName As String, ByVal sourceRow As Long, ByVal cnNo As String, _
                          ByVal invoiceNo As String, ByVal issueName As String, ByVal details As String)
    If logCount >= UBound(logData, 1) Then
        Err.Raise vbObjectError + 1113, "modMovement.AddValidation", "Validation buffer overflow. Increase LOG allocation."
    End If

    logCount = logCount + 1
    logData(logCount, 1) = stageName
    logData(logCount, 2) = sourceRow
    logData(logCount, 3) = cnNo
    logData(logCount, 4) = invoiceNo
    logData(logCount, 5) = issueName
    logData(logCount, 6) = details
    logData(logCount, 7) = Now
End Sub
