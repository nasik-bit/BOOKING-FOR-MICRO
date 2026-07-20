# REVIEW.md

Complete code review of all changes made during the VBA project audit.
Scope: `VBA/modMaster.bas`, `VBA/modStock.bas`, `VBA/modDateWise.bas`,
`VBA/modUtility.bas`. No other files were modified. Nothing in this
review has been committed.

---

## 1. Complete code diff

### VBA/modMaster.bas

```diff
diff --git a/VBA/modMaster.bas b/VBA/modMaster.bas
index 8f2cf57..7188c56 100644
--- a/VBA/modMaster.bas
+++ b/VBA/modMaster.bas
@@ -9,26 +9,61 @@ Public Sub LoadMaster()
     Dim ws As Worksheet
     Dim r As Long, lastRow As Long
 
+    Dim colItemCode As Long, colIgnore As Long
+    Dim colPackingType As Long, colActive As Long
+    Dim colLocationCode As Long
+
     Set gPacking = CreateObject("Scripting.Dictionary")
     Set gIgnoreItems = CreateObject("Scripting.Dictionary")
     Set gLocations = CreateObject("Scripting.Dictionary")
 
     Set ws = ThisWorkbook.Worksheets("MASTER")
 
-    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
-    For r = 2 To lastRow
-        If Trim(ws.Cells(r, "A").Value) <> "" Then gPacking(UCase$(Trim(ws.Cells(r, "A").Value))) = True
-    Next r
+    colItemCode = GetColumn(ws, "ItemCode")
+    colIgnore = GetColumn(ws, "Ignore")
+    colPackingType = GetColumn(ws, "PackingType")
+    colActive = GetColumn(ws, "Active")
+    colLocationCode = GetColumn(ws, "LocationCode")
+
+    '-------------------------------------------------------
+    ' Ignored Item Codes: ItemCode rows flagged Ignore = "Y"
+    '-------------------------------------------------------
+    If colItemCode > 0 And colIgnore > 0 Then
+        lastRow = ws.Cells(ws.Rows.Count, colItemCode).End(xlUp).Row
+        For r = 2 To lastRow
+            If Trim(ws.Cells(r, colItemCode).Value) <> "" Then
+                If UCase$(Trim(ws.Cells(r, colIgnore).Value)) = "Y" Then
+                    gIgnoreItems(UCase$(Trim(ws.Cells(r, colItemCode).Value))) = True
+                End If
+            End If
+        Next r
+    End If
 
-    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
-    For r = 2 To lastRow
-        If Trim(ws.Cells(r, "C").Value) <> "" Then gIgnoreItems(UCase$(Trim(ws.Cells(r, "C").Value))) = True
-    Next r
+    '-------------------------------------------------------
+    ' Valid Packing Types: PackingType rows flagged Active = "Y"
+    '-------------------------------------------------------
+    If colPackingType > 0 And colActive > 0 Then
+        lastRow = ws.Cells(ws.Rows.Count, colPackingType).End(xlUp).Row
+        For r = 2 To lastRow
+            If Trim(ws.Cells(r, colPackingType).Value) <> "" Then
+                If UCase$(Trim(ws.Cells(r, colActive).Value)) = "Y" Then
+                    gPacking(UCase$(Trim(ws.Cells(r, colPackingType).Value))) = True
+                End If
+            End If
+        Next r
+    End If
 
-    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
-    For r = 2 To lastRow
-        If Trim(ws.Cells(r, "E").Value) <> "" Then gLocations(UCase$(Trim(ws.Cells(r, "E").Value))) = True
-    Next r
+    '-------------------------------------------------------
+    ' Valid Locations: LocationCode column
+    '-------------------------------------------------------
+    If colLocationCode > 0 Then
+        lastRow = ws.Cells(ws.Rows.Count, colLocationCode).End(xlUp).Row
+        For r = 2 To lastRow
+            If Trim(ws.Cells(r, colLocationCode).Value) <> "" Then
+                gLocations(UCase$(Trim(ws.Cells(r, colLocationCode).Value))) = True
+            End If
+        Next r
+    End If
 End Sub
 
 Public Function IsValidPacking(ByVal PackingType As String) As Boolean
```

### VBA/modStock.bas

```diff
diff --git a/VBA/modStock.bas b/VBA/modStock.bas
index c00f100..be34b7e 100644
--- a/VBA/modStock.bas
+++ b/VBA/modStock.bas
@@ -7,13 +7,6 @@ Public Sub GenerateCurrentStock()
 
     Const LOC_COUNT As Long = 4
     Const OUT_ROW_COUNT As Long = 5
-    Const MOV_COL_DATE As Long = 1
-    Const MOV_COL_LOCATION As Long = 7
-    Const MOV_COL_QTY As Long = 8
-    Const OPEN_COL_DATE As Long = 1
-    Const OPEN_COL_TYPE As Long = 2
-    Const OPEN_COL_LOCATION As Long = 3
-    Const OPEN_COL_QTY As Long = 4
 
     Dim wsMov As Worksheet
     Dim wsOpen As Worksheet
@@ -40,21 +33,52 @@ Public Sub GenerateCurrentStock()
     Dim moveArr As Variant
     Dim i As Long
 
+    Dim MOV_COL_DATE As Long
+    Dim MOV_COL_LOCATION As Long
+    Dim MOV_COL_QTY As Long
+    Dim OPEN_COL_DATE As Long
+    Dim OPEN_COL_TYPE As Long
+    Dim OPEN_COL_LOCATION As Long
+    Dim OPEN_COL_QTY As Long
+    Dim movReadCols As Long
+    Dim openReadCols As Long
+
     Set wsMov = ThisWorkbook.Worksheets(SHEET_MOVEMENT)
     Set wsOpen = ThisWorkbook.Worksheets(SHEET_OPENING)
     Set wsStock = CreateSheet(SHEET_STOCK)
     earliestOpenDate = STOCK_MAX_LONG
 
+    MOV_COL_DATE = GetColumn(wsMov, "Event Date")
+    MOV_COL_LOCATION = GetColumn(wsMov, "Location")
+    MOV_COL_QTY = GetColumn(wsMov, "Packets")
+    movReadCols = WorksheetFunction.Max(MOV_COL_DATE, MOV_COL_LOCATION, MOV_COL_QTY)
+
+    OPEN_COL_DATE = GetColumn(wsOpen, "Date")
+    OPEN_COL_TYPE = GetColumn(wsOpen, "Transaction Type")
+    OPEN_COL_LOCATION = GetColumn(wsOpen, "Location")
+    OPEN_COL_QTY = GetColumn(wsOpen, "Qty")
+    openReadCols = WorksheetFunction.Max(OPEN_COL_DATE, OPEN_COL_TYPE, OPEN_COL_LOCATION, OPEN_COL_QTY)
+
+    If MOV_COL_DATE = 0 Or MOV_COL_LOCATION = 0 Or MOV_COL_QTY = 0 Then
+        Err.Raise vbObjectError + 1, "modStock.GenerateCurrentStock", _
+            "MOVEMENT sheet is missing one of the required headers: Event Date, Location, Packets."
+    End If
+
+    If OPEN_COL_DATE = 0 Or OPEN_COL_TYPE = 0 Or OPEN_COL_LOCATION = 0 Or OPEN_COL_QTY = 0 Then
+        Err.Raise vbObjectError + 2, "modStock.GenerateCurrentStock", _
+            "OPENING_BALANCE sheet is missing one of the required headers: Date, Transaction Type, Location, Qty."
+    End If
+
     wsStock.Cells.Clear
     wsStock.Range("A1:B1").Value = Array("Location", "Current Stock")
 
-    openLast = LastRow(wsOpen, 1)
+    openLast = LastRow(wsOpen, OPEN_COL_DATE)
     If openLast >= 2 Then
-        openData = wsOpen.Range("A2:D" & openLast).Value2
+        openData = wsOpen.Range(wsOpen.Cells(2, 1), wsOpen.Cells(openLast, openReadCols)).Value2
 
         For r = 1 To UBound(openData, 1)
             idx = LocationIndex(openData(r, OPEN_COL_LOCATION))
-            If idx > 0 And IsDate(openData(r, OPEN_COL_DATE)) Then
+            If idx > 0 And IsValidDateValue(openData(r, OPEN_COL_DATE)) Then
                 dKey = CLng(CDate(openData(r, OPEN_COL_DATE)))
                 If Not hasOpeningDate Or dKey < earliestOpenDate Then
                     earliestOpenDate = dKey
@@ -66,7 +90,7 @@ Public Sub GenerateCurrentStock()
         For r = 1 To UBound(openData, 1)
             idx = LocationIndex(openData(r, OPEN_COL_LOCATION))
             If idx > 0 And hasOpeningDate Then
-                If IsDate(openData(r, OPEN_COL_DATE)) Then
+                If IsValidDateValue(openData(r, OPEN_COL_DATE)) Then
                     dKey = CLng(CDate(openData(r, OPEN_COL_DATE)))
                 Else
                     dKey = 0
@@ -79,13 +103,13 @@ Public Sub GenerateCurrentStock()
         Next r
     End If
 
-    movLast = LastRow(wsMov, 1)
+    movLast = LastRow(wsMov, MOV_COL_DATE)
     If movLast >= 2 Then
-        movData = wsMov.Range("A2:H" & movLast).Value2
+        movData = wsMov.Range(wsMov.Cells(2, 1), wsMov.Cells(movLast, movReadCols)).Value2
         Set moveDict = CreateObject("Scripting.Dictionary")
 
         For r = 1 To UBound(movData, 1)
-            If IsDate(movData(r, MOV_COL_DATE)) Then
+            If IsValidDateValue(movData(r, MOV_COL_DATE)) Then
                 idx = LocationIndex(movData(r, MOV_COL_LOCATION))
                 If idx > 0 Then
                     dKey = CLng(CDate(movData(r, MOV_COL_DATE)))
```

### VBA/modDateWise.bas

```diff
diff --git a/VBA/modDateWise.bas b/VBA/modDateWise.bas
index 153d832..6379f71 100644
--- a/VBA/modDateWise.bas
+++ b/VBA/modDateWise.bas
@@ -41,16 +41,16 @@ Option Explicit
 ' --- Error-log sheet name ------------------------------------------------
 Private Const SHEET_ERROR_LOG As String = "ERROR_LOG"
 
-' --- MOVEMENT column offsets (relative to A2:H bulk read) ----------------
-Private Const MOV_DATE As Long = 1   ' Col A  Event Date
-Private Const MOV_LOC  As Long = 7   ' Col G  Location
-Private Const MOV_QTY  As Long = 8   ' Col H  Packets (positive = IN, negative = OUT)
+' --- MOVEMENT column offsets (resolved at runtime via GetColumn) ---------
+Private MOV_DATE As Long   ' Col "Event Date"
+Private MOV_LOC  As Long   ' Col "Location"
+Private MOV_QTY  As Long   ' Col "Packets" (positive = IN, negative = OUT)
 
-' --- OPENING_BALANCE column offsets (relative to A2:D bulk read) ---------
-Private Const OPEN_DATE As Long = 1  ' Col A  Date
-Private Const OPEN_TYPE As Long = 2  ' Col B  Transaction Type
-Private Const OPEN_LOC  As Long = 3  ' Col C  Location
-Private Const OPEN_QTY  As Long = 4  ' Col D  Quantity
+' --- OPENING_BALANCE column offsets (resolved at runtime via GetColumn) --
+Private OPEN_DATE As Long  ' Col "Date"
+Private OPEN_TYPE As Long  ' Col "Transaction Type"
+Private OPEN_LOC  As Long  ' Col "Location"
+Private OPEN_QTY  As Long  ' Col "Qty"
 
 ' --- Location constants (1-based) ----------------------------------------
 Private Const LOC_COUNT As Long = 4  ' M&M=1  NASHIK=2  PUNE=3  CK=4
@@ -117,6 +117,8 @@ Public Sub GenerateDateWiseStock()
 
     Dim openLast As Long
     Dim movLast  As Long
+    Dim openReadCols As Long
+    Dim movReadCols  As Long
 
     AppStart
     On Error GoTo ErrHandler
@@ -129,17 +131,41 @@ Public Sub GenerateDateWiseStock()
     Set wsDW   = CreateSheet(SHEET_DATE_WISE)
     earliestOpenDate = DATEWISE_MAX_LONG
 
+    ' ------------------------------------------------------------------
+    ' Resolve column positions dynamically (no hardcoded column numbers)
+    ' ------------------------------------------------------------------
+    MOV_DATE = GetColumn(wsMov, "Event Date")
+    MOV_LOC = GetColumn(wsMov, "Location")
+    MOV_QTY = GetColumn(wsMov, "Packets")
+    movReadCols = WorksheetFunction.Max(MOV_DATE, MOV_LOC, MOV_QTY)
+
+    OPEN_DATE = GetColumn(wsOpen, "Date")
+    OPEN_TYPE = GetColumn(wsOpen, "Transaction Type")
+    OPEN_LOC = GetColumn(wsOpen, "Location")
+    OPEN_QTY = GetColumn(wsOpen, "Qty")
+    openReadCols = WorksheetFunction.Max(OPEN_DATE, OPEN_TYPE, OPEN_LOC, OPEN_QTY)
+
+    If MOV_DATE = 0 Or MOV_LOC = 0 Or MOV_QTY = 0 Then
+        Err.Raise vbObjectError + 1, "modDateWise.GenerateDateWiseStock", _
+            "MOVEMENT sheet is missing one of the required headers: Event Date, Location, Packets."
+    End If
+
+    If OPEN_DATE = 0 Or OPEN_TYPE = 0 Or OPEN_LOC = 0 Or OPEN_QTY = 0 Then
+        Err.Raise vbObjectError + 2, "modDateWise.GenerateDateWiseStock", _
+            "OPENING_BALANCE sheet is missing one of the required headers: Date, Transaction Type, Location, Qty."
+    End If
+
     wsDW.Cells.Clear   ' Wipe content AND formatting from any prior run
 
     ' ==================================================================
     ' Step 1  –  Seed running totals from earliest OPENING_BALANCE date only
     ' ==================================================================
-    openLast = LastRow(wsOpen, 1)
+    openLast = LastRow(wsOpen, OPEN_DATE)
     If openLast >= 2 Then
-        openData = wsOpen.Range("A2:D" & openLast).Value2
+        openData = wsOpen.Range(wsOpen.Cells(2, 1), wsOpen.Cells(openLast, openReadCols)).Value2
         For r = 1 To UBound(openData, 1)
             idx = DW_LocationIndex(openData(r, OPEN_LOC))
-            If idx > 0 And IsDate(openData(r, OPEN_DATE)) Then
+            If idx > 0 And IsValidDateValue(openData(r, OPEN_DATE)) Then
                 dKey = CLng(CDate(openData(r, OPEN_DATE)))
                 If Not hasOpeningDate Or dKey < earliestOpenDate Then
                     earliestOpenDate = dKey
@@ -151,7 +177,7 @@ Public Sub GenerateDateWiseStock()
         For r = 1 To UBound(openData, 1)
             idx = DW_LocationIndex(openData(r, OPEN_LOC))
             If idx > 0 And hasOpeningDate Then
-                If IsDate(openData(r, OPEN_DATE)) Then
+                If IsValidDateValue(openData(r, OPEN_DATE)) Then
                     dKey = CLng(CDate(openData(r, OPEN_DATE)))
                 Else
                     dKey = 0
@@ -172,9 +198,9 @@ Public Sub GenerateDateWiseStock()
     ' ==================================================================
     ' Step 2  –  Read MOVEMENT into bulk array
     ' ==================================================================
-    movLast = LastRow(wsMov, 1)
+    movLast = LastRow(wsMov, MOV_DATE)
     If movLast >= 2 Then
-        movData = wsMov.Range("A2:H" & movLast).Value2
+        movData = wsMov.Range(wsMov.Cells(2, 1), wsMov.Cells(movLast, movReadCols)).Value2
 
         ' ==================================================================
         ' Step 3  –  Build per-date In / Out delta dictionary (O(n) single pass)
@@ -183,7 +209,7 @@ Public Sub GenerateDateWiseStock()
         '   Negative Packets value  →  OUT (slot locOff + 1, stored as Abs)
         ' ==================================================================
         For r = 1 To UBound(movData, 1)
-            If IsDate(movData(r, MOV_DATE)) Then               ' Rule 9: skip blank rows
+            If IsValidDateValue(movData(r, MOV_DATE)) Then               ' Rule 9: skip blank rows
                 idx = DW_LocationIndex(movData(r, MOV_LOC))
                 If idx > 0 Then
                     dKey = CLng(CDate(movData(r, MOV_DATE)))
```

### VBA/modUtility.bas

```diff
diff --git a/VBA/modUtility.bas b/VBA/modUtility.bas
index 3370915..ce15b9b 100644
--- a/VBA/modUtility.bas
+++ b/VBA/modUtility.bas
@@ -137,6 +137,21 @@ Public Function FDate(ByVal v As Variant) As Variant
     End If
 End Function
 
+' IsDate() only recognizes true Date variants or date-formatted strings.
+' Bulk range reads via .Value2 (used for fast array processing) return
+' date-formatted cells as plain Doubles, which IsDate() always evaluates
+' to False, silently skipping every row. Use this helper wherever dates
+' are read via .Value2 arrays instead of calling IsDate() directly.
+Public Function IsValidDateValue(ByVal v As Variant) As Boolean
+    If IsDate(v) Then
+        IsValidDateValue = True
+    ElseIf IsNumeric(v) Then
+        IsValidDateValue = (CDbl(v) > 0)
+    Else
+        IsValidDateValue = False
+    End If
+End Function
+
 Public Function FindDeliveryRow(CNNo As String) As Long
     Dim ws As Worksheet
     Dim f As Range
```

---

## 2. Every new function added

| Function | File | Signature |
|---|---|---|
| `IsValidDateValue` | `VBA/modUtility.bas` | `Public Function IsValidDateValue(ByVal v As Variant) As Boolean` |

This is the **only** new procedure introduced across all four files. No other
function was added. No existing function was renamed, and no procedure
was removed.

---

## 3. Every deleted line

### VBA/modMaster.bas
```
-    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
-    For r = 2 To lastRow
-        If Trim(ws.Cells(r, "A").Value) <> "" Then gPacking(UCase$(Trim(ws.Cells(r, "A").Value))) = True
-    Next r
-    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
-    For r = 2 To lastRow
-        If Trim(ws.Cells(r, "C").Value) <> "" Then gIgnoreItems(UCase$(Trim(ws.Cells(r, "C").Value))) = True
-    Next r
-    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
-    For r = 2 To lastRow
-        If Trim(ws.Cells(r, "E").Value) <> "" Then gLocations(UCase$(Trim(ws.Cells(r, "E").Value))) = True
-    Next r
```

### VBA/modStock.bas
```
-    Const MOV_COL_DATE As Long = 1
-    Const MOV_COL_LOCATION As Long = 7
-    Const MOV_COL_QTY As Long = 8
-    Const OPEN_COL_DATE As Long = 1
-    Const OPEN_COL_TYPE As Long = 2
-    Const OPEN_COL_LOCATION As Long = 3
-    Const OPEN_COL_QTY As Long = 4
```
(remaining removals in this file are 1-for-1 line replacements — see
Section 4, "changed lines" — not pure deletions.)

### VBA/modDateWise.bas
```
-' --- MOVEMENT column offsets (relative to A2:H bulk read) ----------------
-Private Const MOV_DATE As Long = 1   ' Col A  Event Date
-Private Const MOV_LOC  As Long = 7   ' Col G  Location
-Private Const MOV_QTY  As Long = 8   ' Col H  Packets (positive = IN, negative = OUT)
-' --- OPENING_BALANCE column offsets (relative to A2:D bulk read) ---------
-Private Const OPEN_DATE As Long = 1  ' Col A  Date
-Private Const OPEN_TYPE As Long = 2  ' Col B  Transaction Type
-Private Const OPEN_LOC  As Long = 3  ' Col C  Location
-Private Const OPEN_QTY  As Long = 4  ' Col D  Quantity
```

### VBA/modUtility.bas
No lines deleted. Insertion only.

---

## 4. Every changed line

### VBA/modMaster.bas
- Column resolution: `"A"` / `"C"` / `"E"` literal column letters →
  `GetColumn(ws, "ItemCode"/"Ignore"/"PackingType"/"Active"/"LocationCode")`.
- `gPacking(...)` population source: was column A (ItemCode) unconditionally →
  now column D (PackingType) filtered by column E (Active = "Y").
- `gIgnoreItems(...)` population source: was column C (blank/unused) →
  now column A (ItemCode) filtered by column B (Ignore = "Y").
- `gLocations(...)` population source: was column E (Active flag) →
  now column G (LocationCode), unfiltered.

### VBA/modStock.bas
- `Const MOV_COL_DATE/LOCATION/QTY` and `Const OPEN_COL_DATE/TYPE/LOCATION/QTY`
  → converted from module-scope `Const` to `Dim`-declared `Long`, assigned via
  `GetColumn(...)` at runtime instead of literal numbers.
- `openLast = LastRow(wsOpen, 1)` → `openLast = LastRow(wsOpen, OPEN_COL_DATE)`.
- `openData = wsOpen.Range("A2:D" & openLast).Value2` →
  `openData = wsOpen.Range(wsOpen.Cells(2, 1), wsOpen.Cells(openLast, openReadCols)).Value2`.
- `If idx > 0 And IsDate(openData(r, OPEN_COL_DATE)) Then` →
  `If idx > 0 And IsValidDateValue(openData(r, OPEN_COL_DATE)) Then` (2 occurrences).
- `movLast = LastRow(wsMov, 1)` → `movLast = LastRow(wsMov, MOV_COL_DATE)`.
- `movData = wsMov.Range("A2:H" & movLast).Value2` →
  `movData = wsMov.Range(wsMov.Cells(2, 1), wsMov.Cells(movLast, movReadCols)).Value2`.
- `If IsDate(movData(r, MOV_COL_DATE)) Then` →
  `If IsValidDateValue(movData(r, MOV_COL_DATE)) Then`.

### VBA/modDateWise.bas
- `MOV_DATE`/`MOV_LOC`/`MOV_QTY`/`OPEN_DATE`/`OPEN_TYPE`/`OPEN_LOC`/`OPEN_QTY`
  → converted from module-scope `Private Const` (hardcoded numbers) to
  `Private` `Long` variables assigned via `GetColumn(...)` at the top of
  `GenerateDateWiseStock`.
- `openLast = LastRow(wsOpen, 1)` → `openLast = LastRow(wsOpen, OPEN_DATE)`.
- `openData = wsOpen.Range("A2:D" & openLast).Value2` →
  `openData = wsOpen.Range(wsOpen.Cells(2, 1), wsOpen.Cells(openLast, openReadCols)).Value2`.
- `If idx > 0 And IsDate(openData(r, OPEN_DATE)) Then` →
  `If idx > 0 And IsValidDateValue(openData(r, OPEN_DATE)) Then`.
- `If IsDate(openData(r, OPEN_DATE)) Then` →
  `If IsValidDateValue(openData(r, OPEN_DATE)) Then`.
- `movLast = LastRow(wsMov, 1)` → `movLast = LastRow(wsMov, MOV_DATE)`.
- `movData = wsMov.Range("A2:H" & movLast).Value2` →
  `movData = wsMov.Range(wsMov.Cells(2, 1), wsMov.Cells(movLast, movReadCols)).Value2`.
- `If IsDate(movData(r, MOV_DATE)) Then` →
  `If IsValidDateValue(movData(r, MOV_DATE)) Then`.

### VBA/modUtility.bas
No existing lines changed. Insertion only (new function, Section 2).

---

## 5. Explanation of every modification

1. **`modMaster.LoadMaster` column source fix** — The MASTER sheet's real
   layout is `A=ItemCode, B=Ignore, D=PackingType, E=Active, G=LocationCode`.
   The original code read columns A/C/E directly, which populated
   `gPacking` with Item Codes instead of Packing Types. Since
   `modMovement.ProcessBookingRow` checks
   `gPacking.Exists(UCase(Packing))` against real Packing Type values
   ("Empty Bin", "Plastic Bin"), this always failed and **every booking
   row was silently skipped** — no movement was ever generated. Fixed by
   resolving each column by header name via `GetColumn()` and filtering
   on the correct Y/N flag columns. Confirmed with the user before
   applying, per the "do not guess" instruction, since it changes which
   sheet data feeds a core business rule.

2. **`modStock`/`modDateWise` hardcoded column constants → `GetColumn()`** —
   Both modules declared column positions as compile-time `Const`
   numbers, directly violating the explicit instruction to never
   hardcode column numbers and to always use `GetColumn()`. Even though
   the hardcoded numbers happened to match the current sheet layout,
   any future column reorder would silently corrupt the calculation with
   no error. Replaced with `GetColumn()`-resolved variables plus
   `Err.Raise` guards if a required header is missing, so a layout change
   fails loudly instead of silently.

3. **Dynamic range reads** — `"A2:H" & lastRow` / `"A2:D" & lastRow`
   assumed the read columns were always contiguous from column A. Since
   columns are now resolved by name (not necessarily contiguous from A),
   the read range is built from the actual resolved column bounds
   (`wsX.Cells(2,1)` to `wsX.Cells(lastRow, maxCol)`) so the bulk array
   read still captures every column the code needs.

4. **`IsDate()` → `IsValidDateValue()` on bulk-read data** — Both
   `GenerateCurrentStock` and `GenerateDateWiseStock` read data via
   `.Range(...).Value2` for performance, then tested dates with
   `IsDate(...)`. `.Value2` returns date-formatted cells as plain
   `Double` serial numbers (not the `Date` subtype), and `IsDate()` on a
   bare `Double` always returns `False` in VBA — verified empirically via
   a live Excel COM test. This meant every date-based row filter in both
   subs silently rejected all rows: opening balances were never seeded
   and no movement deltas were ever added, so both reports always
   produced zero/empty output regardless of the underlying data. Added
   `modUtility.IsValidDateValue()`, which accepts both true dates/strings
   (`IsDate`) and the numeric doubles produced by `.Value2` reads, and
   replaced all six affected call sites (three in each module). The
   downstream `CDate(...)` conversions were already correct and left
   unchanged — only the broken detection check was replaced.

5. **New helper placement** — `IsValidDateValue` was added to
   `modUtility.bas` (not duplicated in `modStock.bas`/`modDateWise.bas`)
   because it is a generic, sheet-agnostic helper consumed by two
   different modules, consistent with the existing pattern of shared
   helpers (`GetColumn`, `LastRow`, `Nz`, `FDate`, `NextDay`) already
   living in `modUtility.bas`.

No procedure was renamed. No module was renamed. No sheet was renamed.
No report layout was changed. No business rule (CNSK/CPUN routing
logic, prefixes, Y/N semantics) was altered — only the mechanism used to
locate columns and detect valid dates.
