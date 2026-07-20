Attribute VB_Name = "modMaster"
Option Explicit

Public gPacking As Object
Public gIgnoreItems As Object
Public gLocations As Object

Public Sub LoadMaster()
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long

    Dim colItemCode As Long, colIgnore As Long
    Dim colPackingType As Long, colActive As Long
    Dim colLocationCode As Long

    Set gPacking = CreateObject("Scripting.Dictionary")
    Set gIgnoreItems = CreateObject("Scripting.Dictionary")
    Set gLocations = CreateObject("Scripting.Dictionary")

    Set ws = ThisWorkbook.Worksheets("MASTER")

    colItemCode = GetColumn(ws, "ItemCode")
    colIgnore = GetColumn(ws, "Ignore")
    colPackingType = GetColumn(ws, "PackingType")
    colActive = GetColumn(ws, "Active")
    colLocationCode = GetColumn(ws, "LocationCode")

    '-------------------------------------------------------
    ' Ignored Item Codes: ItemCode rows flagged Ignore = "Y"
    '-------------------------------------------------------
    If colItemCode > 0 And colIgnore > 0 Then
        lastRow = ws.Cells(ws.Rows.Count, colItemCode).End(xlUp).Row
        For r = 2 To lastRow
            If Trim(ws.Cells(r, colItemCode).Value) <> "" Then
                If UCase$(Trim(ws.Cells(r, colIgnore).Value)) = "Y" Then
                    gIgnoreItems(UCase$(Trim(ws.Cells(r, colItemCode).Value))) = True
                End If
            End If
        Next r
    End If

    '-------------------------------------------------------
    ' Valid Packing Types: PackingType rows flagged Active = "Y"
    '-------------------------------------------------------
    If colPackingType > 0 And colActive > 0 Then
        lastRow = ws.Cells(ws.Rows.Count, colPackingType).End(xlUp).Row
        For r = 2 To lastRow
            If Trim(ws.Cells(r, colPackingType).Value) <> "" Then
                If UCase$(Trim(ws.Cells(r, colActive).Value)) = "Y" Then
                    gPacking(UCase$(Trim(ws.Cells(r, colPackingType).Value))) = True
                End If
            End If
        Next r
    End If

    '-------------------------------------------------------
    ' Valid Locations: LocationCode column
    '-------------------------------------------------------
    If colLocationCode > 0 Then
        lastRow = ws.Cells(ws.Rows.Count, colLocationCode).End(xlUp).Row
        For r = 2 To lastRow
            If Trim(ws.Cells(r, colLocationCode).Value) <> "" Then
                gLocations(UCase$(Trim(ws.Cells(r, colLocationCode).Value))) = True
            End If
        Next r
    End If
End Sub

Public Function IsValidPacking(ByVal PackingType As String) As Boolean
    If gPacking Is Nothing Then LoadMaster
    IsValidPacking = gPacking.Exists(UCase$(Trim(PackingType)))
End Function

Public Function IsIgnoredItem(ByVal ItemCode As String) As Boolean
    If gIgnoreItems Is Nothing Then LoadMaster
    IsIgnoredItem = gIgnoreItems.Exists(UCase$(Trim(ItemCode)))
End Function

Public Function IsValidLocation(ByVal Location As String) As Boolean
    If gLocations Is Nothing Then LoadMaster
    IsValidLocation = gLocations.Exists(UCase$(Trim(Location)))
End Function
