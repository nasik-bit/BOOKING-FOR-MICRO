Attribute VB_Name = "modMaster"
Option Explicit

Public gPacking As Object
Public gIgnoreItems As Object
Public gLocations As Object

Public Sub LoadMaster()
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long

    Set gPacking = CreateObject("Scripting.Dictionary")
    Set gIgnoreItems = CreateObject("Scripting.Dictionary")
    Set gLocations = CreateObject("Scripting.Dictionary")

    Set ws = ThisWorkbook.Worksheets("MASTER")

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For r = 2 To lastRow
        If Trim(ws.Cells(r, "A").Value) <> "" Then gPacking(UCase$(Trim(ws.Cells(r, "A").Value))) = True
    Next r

    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    For r = 2 To lastRow
        If Trim(ws.Cells(r, "C").Value) <> "" Then gIgnoreItems(UCase$(Trim(ws.Cells(r, "C").Value))) = True
    Next r

    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    For r = 2 To lastRow
        If Trim(ws.Cells(r, "E").Value) <> "" Then gLocations(UCase$(Trim(ws.Cells(r, "E").Value))) = True
    Next r
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
