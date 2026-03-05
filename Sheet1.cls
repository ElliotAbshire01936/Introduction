VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' =====================================================================
' Sheet1.cls  –  Worksheet module for barcode / inventory management
'
' PERFORMANCE NOTES (vs. original code):
'   1. SelectionChange now exits early when outside the data area,
'      avoiding a full workbook recalculation for every click.
'   2. Worksheet_Change routes to focused helpers instead of
'      monolithic logic, keeping each handler small.
'   3. Barcode deduplication uses a Scripting.Dictionary (O(1) lookup)
'      built once per change event, eliminating repeated CountIf calls.
'   4. Override look-ups use Range.Find instead of CountIf / loops.
'   5. SortDataWithoutHeader finds the actual last data row before
'      sorting so it never processes empty rows.
'   6. Application.EnableEvents is always restored in error handlers;
'      a module-level flag (mEventGuard) prevents re-entrant calls.
' =====================================================================

' ----- Module-level constants ----------------------------------------

Private Const MAX_ROW           As Long = 1000  ' Data area ceiling
Private Const COL_BARCODE       As Long = 2     ' Column B
Private Const COL_OVERRIDE_L    As Long = 12    ' Column L
Private Const COL_OVERRIDE_M    As Long = 13    ' Column M
Private Const OVERRIDE_FIRST_ROW As Long = 6   ' Overrides start at row 6
Private Const OVR_KEY_COL       As Long = 1     ' OVERRIDES!A  – barcode key
Private Const OVR_L_COL         As Long = 2     ' OVERRIDES!B  – L value
Private Const OVR_M_COL         As Long = 3     ' OVERRIDES!C  – M value

' Re-entrancy guard: set True while a Change handler is running so that
' any programmatic cell writes inside the helper don't fire a new event.
Private mEventGuard As Boolean

' =====================================================================
' Worksheet_SelectionChange
'
' PERF: Original code called Application.Calculate (full workbook) for
' every selection change anywhere on the sheet.  Now we exit early when
' the selection is outside the data area (rows 1–MAX_ROW, cols A–M),
' and we recalculate only Me.UsedRange rather than the whole workbook.
' =====================================================================
Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    ' Exit immediately when the selection is completely outside the data
    ' block – avoids any recalculation overhead for those clicks.
    If Target.Row > MAX_ROW Then Exit Sub
    If Target.Column > COL_OVERRIDE_M Then Exit Sub

    ' Recalculate only the used range on this sheet (not the whole
    ' workbook) – sufficient for formula cells that depend on selection.
    Me.UsedRange.Calculate

End Sub

' =====================================================================
' Worksheet_Change
'
' Central router.  Determines which area(s) of the sheet were affected
' and delegates to focused helpers.  Uses mEventGuard to prevent the
' helper's programmatic writes from firing another Change event.
' =====================================================================
Private Sub Worksheet_Change(ByVal Target As Range)

    ' Re-entrancy guard: bail out if we triggered this ourselves.
    If mEventGuard Then Exit Sub

    Dim rBarcode  As Range  ' B2:B<MAX_ROW>
    Dim rOverride As Range  ' L6:M<MAX_ROW>
    Dim rHit      As Range

    Set rBarcode  = Me.Range("B2:B" & MAX_ROW)
    Set rOverride = Me.Range("L" & OVERRIDE_FIRST_ROW & ":M" & MAX_ROW)

    ' --- Barcode column (B) ------------------------------------------
    Set rHit = Intersect(Target, rBarcode)
    If Not rHit Is Nothing Then
        Call HandleBarcodeChange(rHit)
    End If

    ' --- Override columns (L : M) -------------------------------------
    Set rHit = Intersect(Target, rOverride)
    If Not rHit Is Nothing Then
        Dim wsOvr As Worksheet
        On Error Resume Next
        Set wsOvr = ThisWorkbook.Sheets("OVERRIDES")
        On Error GoTo 0
        If Not wsOvr Is Nothing Then
            Call HandleOverrideChange(rHit, wsOvr)
        End If
    End If

End Sub

' =====================================================================
' HandleBarcodeChange
'
' For every cell in the changed range (handles single edits AND large
' pastes with one pass):
'   1. Normalises the raw value via NormalizeBarcode().
'   2. Stores it back as text (NumberFormat = "@") so leading zeros and
'      long numeric codes are preserved.
'   3. Checks for duplicates against all other barcodes in column B
'      using a Scripting.Dictionary built once (O(1) per lookup) rather
'      than a CountIf call per cell.
'
' PERF vs. original: eliminated N CountIf calls for N pasted cells;
' replaced with a single O(n) scan to build a Dictionary, then O(1)
' lookups.
' =====================================================================
Private Sub HandleBarcodeChange(ByVal changedRange As Range)

    Dim oDict        As Object   ' Scripting.Dictionary
    Dim rAllBarcodes As Range
    Dim cell         As Range
    Dim raw          As String
    Dim normalised   As String
    Dim dupFound     As Boolean

    mEventGuard = True
    Application.EnableEvents = False

    On Error GoTo CleanupBarcode

    ' Build a Dictionary of every existing barcode in B2:B<MAX_ROW>
    ' that is NOT part of this change (so we don't flag self-duplicates
    ' during normalisation rewrites).
    Set oDict = CreateObject("Scripting.Dictionary")
    oDict.CompareMode = 1  ' vbTextCompare – case-insensitive

    Set rAllBarcodes = Me.Range("B2:B" & MAX_ROW)
    For Each cell In rAllBarcodes.Cells
        If Intersect(cell, changedRange) Is Nothing Then
            raw = Trim(CStr(cell.Value))
            If raw <> "" Then
                normalised = NormalizeBarcode(raw)
                If normalised <> "" Then
                    oDict(normalised) = cell.Row  ' value = row for diagnostics
                End If
            End If
        End If
    Next cell

    ' Process each changed cell.
    dupFound = False
    For Each cell In changedRange.Cells
        raw = Trim(CStr(cell.Value))

        If raw = "" Then
            ' Cleared cell – keep text format so future entries stay text.
            cell.NumberFormat = "@"
        Else
            normalised = NormalizeBarcode(raw)

            If oDict.Exists(normalised) Then
                ' Duplicate: clear the cell.
                cell.ClearContents
                cell.NumberFormat = "@"
                dupFound = True
            Else
                ' Accepted: store normalised value as text.
                cell.NumberFormat = "@"
                cell.Value = normalised
                ' Record in Dictionary so intra-paste duplicates are caught.
                oDict(normalised) = cell.Row
            End If
        End If
    Next cell

CleanupBarcode:
    Application.EnableEvents = True
    mEventGuard = False

    If Err.Number <> 0 Then
        MsgBox "Error in barcode handling: " & Err.Description, _
               vbCritical, "Barcode Error"
    ElseIf dupFound Then
        MsgBox "One or more duplicate barcodes were detected and cleared.", _
               vbExclamation, "Duplicate Barcode"
    End If

End Sub

' =====================================================================
' HandleOverrideChange
'
' When the user types, pastes, or clears cells in L6:M1000:
'   – Reads the normalised barcode from column A of the same row.
'   – Finds or creates a matching row in the OVERRIDES sheet (keyed on
'     col A of OVERRIDES).
'   – Writes L-column values → OVERRIDES col B,
'           M-column values → OVERRIDES col C.
'   – After a multi-cell paste, calls RestoreOverrideFormulas() to put
'     back standard formulas in cells that don't have a manual override.
'
' PERF vs. original: uses Find() (one pass) instead of CountIf per row.
' =====================================================================
Private Sub HandleOverrideChange(ByVal changedRange As Range, _
                                 ByVal wsOvr As Worksheet)

    Dim cell       As Range
    Dim barcodeKey As String
    Dim ovrRow     As Long
    Dim isMulti    As Boolean

    isMulti = (changedRange.Cells.Count > 1)

    mEventGuard = True
    Application.EnableEvents = False

    On Error GoTo CleanupOverride

    For Each cell In changedRange.Cells
        ' Barcode key from column A of this row.
        barcodeKey = NormalizeBarcode(CStr(Me.Cells(cell.Row, 1).Value))
        If barcodeKey = "" Then GoTo NextCell

        ovrRow = FindOrCreateOverrideRow(wsOvr, barcodeKey)
        If ovrRow = 0 Then GoTo NextCell

        Select Case cell.Column
            Case COL_OVERRIDE_L
                wsOvr.Cells(ovrRow, OVR_L_COL).Value = cell.Value
            Case COL_OVERRIDE_M
                wsOvr.Cells(ovrRow, OVR_M_COL).Value = cell.Value
        End Select

NextCell:
    Next cell

    ' After a multi-cell paste, restore standard lookup formulas in any
    ' L/M cells that do NOT have a row in the OVERRIDES sheet (i.e. the
    ' paste accidentally overwrote formula rows with static values).
    If isMulti Then
        Call RestoreOverrideFormulas(changedRange, wsOvr)
    End If

CleanupOverride:
    Application.EnableEvents = True
    mEventGuard = False

    If Err.Number <> 0 Then
        MsgBox "Error in override handling: " & Err.Description, _
               vbCritical, "Override Error"
    End If

End Sub

' =====================================================================
' FindOrCreateOverrideRow
'
' Searches OVERRIDES column A for barcodeKey using Range.Find (fast;
' avoids a full column scan or CountIf).  If not found, appends a new
' row at the bottom and returns its row number.
' Returns 0 on failure.
' =====================================================================
Private Function FindOrCreateOverrideRow(ByVal wsOvr As Worksheet, _
                                         ByVal barcodeKey As String) As Long
    Dim foundCell As Range
    Dim lastRow   As Long

    On Error GoTo FindFail

    ' Find existing entry – faster than a CountIf or manual loop
    ' because Range.Find stops at the first match.
    Set foundCell = wsOvr.Columns(OVR_KEY_COL).Find( _
        What:=barcodeKey, _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        MatchCase:=False)

    If Not foundCell Is Nothing Then
        FindOrCreateOverrideRow = foundCell.Row
        Exit Function
    End If

    ' Not found: append a new row.
    lastRow = wsOvr.Cells(wsOvr.Rows.Count, OVR_KEY_COL).End(xlUp).Row
    If wsOvr.Cells(lastRow, OVR_KEY_COL).Value <> "" Then
        lastRow = lastRow + 1
    End If
    wsOvr.Cells(lastRow, OVR_KEY_COL).Value = barcodeKey
    FindOrCreateOverrideRow = lastRow
    Exit Function

FindFail:
    FindOrCreateOverrideRow = 0
End Function

' =====================================================================
' RestoreOverrideFormulas
'
' After a multi-cell paste into L:M, identifies cells whose barcode row
' does NOT have a manual entry in OVERRIDES, and re-applies the default
' lookup formula so static pasted values don't permanently replace them.
'
' The formula templates in GetDefaultFormulaL / GetDefaultFormulaM can
' be updated to match the workbook's actual lookup logic.
' =====================================================================
Private Sub RestoreOverrideFormulas(ByVal pastedRange As Range, _
                                    ByVal wsOvr As Worksheet)
    Dim cell       As Range
    Dim barcodeKey As String
    Dim foundCell  As Range

    For Each cell In pastedRange.Cells
        barcodeKey = NormalizeBarcode(CStr(Me.Cells(cell.Row, 1).Value))
        If barcodeKey = "" Then GoTo NextRestoreCell

        ' Check whether a manual override exists for this barcode.
        Set foundCell = wsOvr.Columns(OVR_KEY_COL).Find( _
            What:=barcodeKey, LookIn:=xlValues, _
            LookAt:=xlWhole, MatchCase:=False)

        ' Only restore the formula if there is NO manual override row,
        ' meaning the paste overwrote a formula cell unintentionally.
        If foundCell Is Nothing Then
            Select Case cell.Column
                Case COL_OVERRIDE_L
                    cell.Formula = GetDefaultFormulaL(cell.Row)
                Case COL_OVERRIDE_M
                    cell.Formula = GetDefaultFormulaM(cell.Row)
            End Select
        End If

NextRestoreCell:
    Next cell
End Sub

' =====================================================================
' GetDefaultFormulaL  /  GetDefaultFormulaM
'
' Returns the standard lookup formula for each override column.
' Edit these two functions to match the workbook's actual logic –
' all other code remains unchanged.
' =====================================================================
Private Function GetDefaultFormulaL(ByVal rowNum As Long) As String
    ' Looks up the barcode (col B) in OVERRIDES A:B → returns col B value.
    ' Bounded range avoids scanning all 1,048,576 rows.
    GetDefaultFormulaL = "=IFERROR(VLOOKUP(B" & rowNum & ",OVERRIDES!$A$1:$B$1000,2,0),"""")"
End Function

Private Function GetDefaultFormulaM(ByVal rowNum As Long) As String
    ' Looks up the barcode (col B) in OVERRIDES A:C → returns col C value.
    ' Bounded range avoids scanning all 1,048,576 rows.
    GetDefaultFormulaM = "=IFERROR(VLOOKUP(B" & rowNum & ",OVERRIDES!$A$1:$C$1000,3,0),"""")"
End Function

' =====================================================================
' NormalizeBarcode
'
' Canonical form: trimmed, upper-cased, retains A–Z / 0–9 / hyphen.
' All other characters are stripped.  Returns "" for blank / all-noise
' input.
'
' Public so it can be called from other modules or tested independently.
' =====================================================================
Public Function NormalizeBarcode(ByVal rawValue As String) As String
    Dim result  As String
    Dim cleaned As String
    Dim i       As Long
    Dim c       As String

    result = UCase(Trim(rawValue))
    cleaned = ""

    For i = 1 To Len(result)
        c = Mid(result, i, 1)
        If (c >= "A" And c <= "Z") Or _
           (c >= "0" And c <= "9") Or _
           c = "-" Then
            cleaned = cleaned & c
        End If
    Next i

    NormalizeBarcode = cleaned
End Function

' =====================================================================
' SortDataWithoutHeader
'
' Sorts the data block (rows 2 – actual last row, columns A–M) by
' column B (barcodes) ascending.
'
' PERF vs. original:
'   – Finds the real last row in column B with End(xlUp) so empty rows
'     below data are never included in the sort range.
'   – Disables ScreenUpdating for the duration of the sort.
'   – Uses .Sort object with explicit range/key references (no
'     CurrentRegion / UsedRange which can expand unexpectedly).
' =====================================================================
Public Sub SortDataWithoutHeader()

    Dim sortRange As Range
    Dim lastRow   As Long

    ' Find the last populated row in the barcode column.
    lastRow = Me.Cells(MAX_ROW, COL_BARCODE).End(xlUp).Row

    ' Nothing to sort if all data rows are empty.
    If lastRow < 2 Then Exit Sub

    ' Explicit range: columns A–M, row 2 to actual last data row.
    Set sortRange = Me.Range(Me.Cells(2, 1), Me.Cells(lastRow, COL_OVERRIDE_M))

    Application.ScreenUpdating = False

    On Error GoTo CleanupSort

    With Me.Sort
        .SortFields.Clear
        .SortFields.Add _
            Key:=Me.Range("B2:B" & lastRow), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal
        .SetRange sortRange
        .Header = xlNo           ' Header row (row 1) excluded from range
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

CleanupSort:
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "Error during sort: " & Err.Description, vbCritical, "Sort Error"
    End If

End Sub
