Attribute VB_Name = "modMovement"
Option Explicit

Public Sub GenerateMovement()

    Dim wsBook As Worksheet
    Dim wsDel As Worksheet
    Dim wsMov As Worksheet

    Dim LastRowBook As Long
    Dim r As Long

    'Load Master
    Call LoadMaster

    Set wsBook = ws("BOOKING_DATA")
    Set wsDel = ws("DELIVERY")
    Set wsMov = CreateSheet("MOVEMENT")

    ClearData wsMov

    'Movement Header
    wsMov.Range("A1:O1").Value = Array( _
        "Event Date", _
        "CN No", _
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
        "Remarks", _
        "Flow")

    LastRowBook = lastRow(wsBook, GetColumn(wsBook, "Consignment Number"))

    For r = 2 To LastRowBook

        ProcessBookingRow wsBook, wsDel, wsMov, r

    Next r

    MsgBox "Movement Generated Successfully.", vbInformation

End Sub
     
Private Sub ProcessBookingRow(wsBook As Worksheet, _
                              wsDel As Worksheet, _
                              wsMov As Worksheet, _
                              ByVal rw As Long)

    Dim Prefix As String
    Dim Packing As String
    Dim ItemCode As String

    Prefix = GetPrefix(wsBook.Cells(rw, GetColumn(wsBook, "Consignment Number")).Value)

    Packing = Trim(wsBook.Cells(rw, GetColumn(wsBook, "Packing Type")).Value)

    ItemCode = Trim(wsBook.Cells(rw, GetColumn(wsBook, "Item Code")).Value)

    '-------------------------
    'Ignore Packing
    '-------------------------

    If Not gPacking.Exists(UCase(Packing)) Then Exit Sub

    '-------------------------
    'Ignore Item Code
    '-------------------------

    If gIgnoreItems.Exists(ItemCode) Then Exit Sub

    '-------------------------
    'Ignore POD
    '-------------------------

    If UCase(wsBook.Cells(rw, GetColumn(wsBook, "CN Status")).Value) = "POD" Then Exit Sub

    Select Case Prefix

        Case "CNSK"

            Call ProcessCNSK(wsBook, wsDel, wsMov, rw)

        Case "CPUN"

            Call ProcessCPUN(wsBook, wsDel, wsMov, rw)

    End Select

End Sub


Private Sub WriteMovement( _
    ws As Worksheet, _
    ByVal EventDate As Variant, _
    ByVal CNNo As String, _
    ByVal Prefix As String, _
    ByVal EventName As String, _
    ByVal Source As String, _
    ByVal Destination As String, _
    ByVal Location As String, _
    ByVal Packets As Double, _
    ByVal PackingType As String, _
    ByVal ItemCode As String, _
    ByVal Consignor As String, _
    ByVal Consignee As String, _
    ByVal RefNo As String, _
    ByVal Remarks As String, _
    ByVal Flow As String)

    Dim nxt As Long

    nxt = lastRow(ws, 1) + 1

    ws.Cells(nxt, 1).Value = EventDate
    ws.Cells(nxt, 2).Value = CNNo
    ws.Cells(nxt, 3).Value = Prefix
    ws.Cells(nxt, 4).Value = EventName
    ws.Cells(nxt, 5).Value = Source
    ws.Cells(nxt, 6).Value = Destination
    ws.Cells(nxt, 7).Value = Location
    ws.Cells(nxt, 8).Value = Packets
    ws.Cells(nxt, 9).Value = PackingType
    ws.Cells(nxt, 10).Value = ItemCode
    ws.Cells(nxt, 11).Value = Consignor
    ws.Cells(nxt, 12).Value = Consignee
    ws.Cells(nxt, 13).Value = RefNo
    ws.Cells(nxt, 14).Value = Remarks
    ws.Cells(nxt, 15).Value = Flow

End Sub

Private Sub ProcessCNSK( _
        wsBook As Worksheet, _
        wsDel As Worksheet, _
        wsMov As Worksheet, _
        ByVal rw As Long)

    Dim CN As String
    Dim CNDate As Variant
    Dim TNDate As Variant
    Dim DNDate As Variant

    Dim Qty As Double

    Dim Packing As String
    Dim ItemCode As String

    Dim Source As String
    Dim Destination As String
    Dim Consignor As String
    Dim Consignee As String

    CN = wsBook.Cells(rw, GetColumn(wsBook, "Consignment Number")).Value
    CNDate = wsBook.Cells(rw, GetColumn(wsBook, "Consignment Date")).Value
    TNDate = wsBook.Cells(rw, GetColumn(wsBook, "Transport Note Date")).Value
    DNDate = wsBook.Cells(rw, GetColumn(wsBook, "DN Date")).Value

    Qty = Nz(wsBook.Cells(rw, GetColumn(wsBook, "Packets")).Value)

    Packing = wsBook.Cells(rw, GetColumn(wsBook, "Packing Type")).Value
    ItemCode = wsBook.Cells(rw, GetColumn(wsBook, "Item Code")).Value

    Source = Trim(wsBook.Cells(rw, GetColumn(wsBook, "Source")).Value)
    Destination = Trim(wsBook.Cells(rw, GetColumn(wsBook, "Destination")).Value)
    Consignor = Trim(wsBook.Cells(rw, GetColumn(wsBook, "Consignor")).Value)
    Consignee = Trim(wsBook.Cells(rw, GetColumn(wsBook, "Consignee")).Value)

    '==========================
    ' BOOKING
    ' M&M (-)
    '==========================
    AddMovementEntry wsMov, _
        CNDate, _
        CN, _
        "CNSK", _
        "BOOKING", _
        Source, _
        Destination, _
        "M&M", _
        Qty * -1, _
        Packing, _
        ItemCode, _
        Consignor, _
        Consignee, _
        "Booking"

    '==========================
    ' TN +1
    ' Pune (+)
    '==========================
    If IsDate(TNDate) Then

        AddMovementEntry wsMov, _
            NextDay(TNDate), _
            CN, _
            "CNSK", _
            "TN+1", _
            Source, _
            Destination, _
            "PUNE", _
            Qty, _
            Packing, _
            ItemCode, _
            Consignor, _
            Consignee, _
            "Transport Note"

    End If

    '==========================
    ' DN
    ' Pune (-)
    '==========================
    If IsDate(DNDate) Then

        AddMovementEntry wsMov, _
            DNDate, _
            CN, _
            "CNSK", _
            "DN", _
            Source, _
            Destination, _
            "PUNE", _
            Qty * -1, _
            Packing, _
            ItemCode, _
            Consignor, _
            Consignee, _
            "Delivery Note"

        '==========================
        ' DN
        ' CK (+)
        '==========================
        AddMovementEntry wsMov, _
            DNDate, _
            CN, _
            "CNSK", _
            "DN", _
            Source, _
            Destination, _
            "CK", _
            Qty, _
            Packing, _
            ItemCode, _
            Consignor, _
            Consignee, _
            "Delivery Note"

    End If

End Sub


'====================================================
' Process CPUN Movement
'====================================================
Private Sub ProcessCPUN( _
        wsBook As Worksheet, _
        wsDel As Worksheet, _
        wsMov As Worksheet, _
        ByVal rw As Long)

    Dim CN As String
    Dim CNDate As Variant
    Dim TNDate As Variant
    Dim DNDate As Variant

    Dim Qty As Double

    Dim Packing As String
    Dim ItemCode As String

    Dim Source As String
    Dim Destination As String
    Dim Consignor As String
    Dim Consignee As String

    CN = wsBook.Cells(rw, GetColumn(wsBook, "Consignment Number")).Value
    CNDate = wsBook.Cells(rw, GetColumn(wsBook, "Consignment Date")).Value
    TNDate = wsBook.Cells(rw, GetColumn(wsBook, "Transport Note Date")).Value
    DNDate = wsBook.Cells(rw, GetColumn(wsBook, "DN Date")).Value

    Qty = Nz(wsBook.Cells(rw, GetColumn(wsBook, "Packets")).Value)

    Packing = Trim(wsBook.Cells(rw, GetColumn(wsBook, "Packing Type")).Value)
    ItemCode = Trim(wsBook.Cells(rw, GetColumn(wsBook, "Item Code")).Value)

    Source = Trim(wsBook.Cells(rw, GetColumn(wsBook, "Source")).Value)
    Destination = Trim(wsBook.Cells(rw, GetColumn(wsBook, "Destination")).Value)
    Consignor = Trim(wsBook.Cells(rw, GetColumn(wsBook, "Consignor")).Value)
    Consignee = Trim(UCase(wsBook.Cells(rw, GetColumn(wsBook, "Consignee")).Value))

    '=========================================
    ' BOOKING
    ' CK (-)
    '=========================================
    AddMovementEntry wsMov, _
        CNDate, _
        CN, _
        "CPUN", _
        "BOOKING", _
        Source, _
        Destination, _
        "CK", _
        Qty * -1, _
        Packing, _
        ItemCode, _
        Consignor, _
        Consignee, _
        "Booking"

    '=========================================
    ' BOOKING
    ' PUNE (+)
    '=========================================
    AddMovementEntry wsMov, _
        CNDate, _
        CN, _
        "CPUN", _
        "BOOKING", _
        Source, _
        Destination, _
        "PUNE", _
        Qty, _
        Packing, _
        ItemCode, _
        Consignor, _
        Consignee, _
        "Booking"

    '=========================================
    ' TN +1
    ' PUNE (-)
    '=========================================
    If IsDate(TNDate) Then

        AddMovementEntry wsMov, _
            NextDay(TNDate), _
            CN, _
            "CPUN", _
            "TN+1", _
            Source, _
            Destination, _
            "PUNE", _
            Qty * -1, _
            Packing, _
            ItemCode, _
            Consignor, _
            Consignee, _
            "Transport Note"

        '=========================================
        ' TN +1
        ' NASHIK (+)
        '=========================================
        AddMovementEntry wsMov, _
            NextDay(TNDate), _
            CN, _
            "CPUN", _
            "TN+1", _
            Source, _
            Destination, _
            "NASHIK", _
            Qty, _
            Packing, _
            ItemCode, _
            Consignor, _
            Consignee, _
            "Transport Note"

    End If

    '=========================================
    ' DN
    ' NASHIK (-)
    ' M&M (+)
    '=========================================
    If IsDate(DNDate) Then

        If Consignee = "MAHINDRA & MAHINDRA LTD." Then

            AddMovementEntry wsMov, _
                DNDate, _
                CN, _
                "CPUN", _
                "DN", _
                Source, _
                Destination, _
                "NASHIK", _
                Qty * -1, _
                Packing, _
                ItemCode, _
                Consignor, _
                Consignee, _
                "Delivery Note"

            AddMovementEntry wsMov, _
                DNDate, _
                CN, _
                "CPUN", _
                "DN", _
                Source, _
                Destination, _
                "M&M", _
                Qty, _
                Packing, _
                ItemCode, _
                Consignor, _
                Consignee, _
                "Delivery Note"

        End If

    End If

End Sub

'=========================================================
' Find Booking Row by CN Number
'=========================================================
Private Function FindBookingRow(wsBook As Worksheet, CNNo As String) As Long

    Dim LastRw As Long
    Dim r As Long
    Dim colCN As Long

    colCN = GetColumn(wsBook, "Consignment Number")

    If colCN = 0 Then Exit Function

    LastRw = lastRow(wsBook, colCN)

    For r = 2 To LastRw

        If Trim(UCase(wsBook.Cells(r, colCN).Value)) = Trim(UCase(CNNo)) Then
            FindBookingRow = r
            Exit Function
        End If

    Next r

    FindBookingRow = 0

End Function

'=========================================================
' Common Procedure to Add One Movement Entry
'=========================================================
Private Sub AddMovementEntry( _
    wsMov As Worksheet, _
    ByVal EventDate As Variant, _
    ByVal CNNo As String, _
    ByVal Prefix As String, _
    ByVal EventName As String, _
    ByVal Source As String, _
    ByVal Destination As String, _
    ByVal Location As String, _
    ByVal Packets As Double, _
    ByVal PackingType As String, _
    ByVal ItemCode As String, _
    ByVal Consignor As String, _
    ByVal Consignee As String, _
    ByVal Remarks As String)

    WriteMovement wsMov, _
        EventDate, _
        CNNo, _
        Prefix, _
        EventName, _
        Source, _
        Destination, _
        Location, _
        Packets, _
        PackingType, _
        ItemCode, _
        Consignor, _
        Consignee, _
        "", _
        Remarks, _
        Prefix

End Sub




