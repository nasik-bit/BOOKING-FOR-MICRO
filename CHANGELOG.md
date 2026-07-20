# Changelog — VBA Project Audit

Scope: `VBA/` folder only (confirmed with user; the compiled `.xlsm` VBA
project is out of sync with this folder — see "Findings (not fixed)" below).

No procedures, modules, sheets, or report layouts were renamed or
redesigned. All business rules (CNSK/CPUN routing, prefixes, Y/N flags,
etc.) are unchanged. All fixes route column lookups through the existing
`GetColumn()` helper — no hardcoded column numbers were introduced.

Validated by importing the modified modules into a real Excel instance
(COM automation), running `GenerateMovement` → `GenerateCurrentStock` →
`GenerateDateWiseStock` against representative synthetic data, and
manually cross-checking the output against the underlying business rules
(opening balance + net transfers = closing balance, per location and
per date; total stock conserved at 850 across all locations/dates in the
test). Compilation was also confirmed implicitly: VBA will not run any
procedure while the project has a compile error, and all three subs ran
and produced correct results.

## Fixes

### 1. `VBA/modMaster.bas` — `LoadMaster` read the wrong MASTER columns (critical)
`gPacking`, `gIgnoreItems`, and `gLocations` were being loaded from
hardcoded columns "A", "C", "E". On the real MASTER sheet layout, column
A holds `ItemCode`, column C is unused, and column E holds an `Active`
Y/N flag — not `PackingType` or `LocationCode`. As a result `gPacking`
ended up populated with Item Codes instead of Packing Types, so
`ProcessBookingRow`'s `gPacking.Exists(UCase(Packing))` check never
matched a real Packing Type value (e.g. "Empty Bin", "Plastic Bin") and
**every booking row was silently skipped** — no movement was ever
generated.

Fixed by resolving each column via `GetColumn()` against the real
headers (`ItemCode`, `Ignore`, `PackingType`, `Active`, `LocationCode`)
and filtering on the correct Y/N flags:
- `gIgnoreItems` ← ItemCode where Ignore = "Y"
- `gPacking` ← PackingType where Active = "Y"
- `gLocations` ← LocationCode

Confirmed with the user before applying (foundational business-rule
assumption).

### 2. `VBA/modStock.bas` and `VBA/modDateWise.bas` — hardcoded column offsets
Both modules declared `Const MOV_COL_*` / `Const OPEN_COL_*` (and
`Const MOV_DATE/MOV_LOC/MOV_QTY/OPEN_DATE/OPEN_TYPE/OPEN_LOC/OPEN_QTY` in
modDateWise) as fixed column numbers instead of using `GetColumn()`,
directly against the explicit instruction to never hardcode column
numbers. Replaced with `GetColumn()` lookups by header name
(`Event Date`, `Location`, `Packets` on MOVEMENT; `Date`,
`Transaction Type`, `Location`, `Qty` on OPENING_BALANCE), each resolved
once at the top of the sub. Added `Err.Raise` guards if a required
header is missing, and changed the bulk-read ranges from hardcoded
`"A2:H" & lastRow` / `"A2:D" & lastRow` strings to ranges sized from the
actual resolved columns.

### 3. `VBA/modStock.bas` and `VBA/modDateWise.bas` — `IsDate()` on bulk-read dates always False (critical, pre-existing)
Both modules read MOVEMENT/OPENING_BALANCE data in bulk via
`.Range(...).Value2` for performance, then tested each row's date with
`IsDate(...)`. `.Value2` returns date-formatted cells as plain `Double`
serial numbers, not the `Date` subtype — and `IsDate()` on a bare
`Double` always returns `False` in VBA (verified empirically). This
meant every date-based row filter in `GenerateCurrentStock` and
`GenerateDateWiseStock` silently rejected every row: opening balances
were never picked up and no movement deltas were ever added, so both
reports always produced zero/empty output regardless of the data.

Added a shared helper, `modUtility.IsValidDateValue()`, that accepts
both true dates/date-strings and the numeric doubles produced by
`.Value2` reads, and replaced the six affected `IsDate(...)` calls (3 in
each module) with it. `CDate()` conversions further down were already
correct and unchanged — only the detection check was broken.

## Findings (not fixed — flagged per user's decision to scope this audit to `VBA/`)
- The compiled VBA project embedded in `BOOKING FOR MICRO.xlsm` is out
  of sync with the `VBA/` folder: `modDateWise` and `modDashboard` are
  empty stubs in the workbook, `modStock` is far less complete than the
  repo version, a `modMain` module exists in the workbook with no
  corresponding repo file, and stray untracked duplicate modules
  `modMaster1` / `modMovement1` exist in the workbook. The workbook's
  VBA project should be re-imported from `VBA/` before the fixes here
  take effect in production.
- Root-level `modMaster_V2.bas`, `modMovement.bas`, and
  `modUtility (2).bas` are older, stale duplicates of files in `VBA/`
  and are not referenced by the audit scope.
- `VBA/modMovement.bas` and `VBA/modUtility.bas` were read in full;
  no compile, runtime, or logic errors were found in either.
