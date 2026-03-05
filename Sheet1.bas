Attribute VB_Name = "Sheet1"
' =============================================================================
' Sheet1 – Worksheet event handlers
'
' Responsibilities
'   • Column B  – Normalise barcodes (trim + uppercase) on entry.
'   • Column B  – Reject duplicate barcodes immediately.
'   • Columns L/M – Write manual overrides to the OVERRIDES sheet, keyed by
'                   the barcode in column B of the same row.
'   • General   – Restore column formulas that may have been destroyed by a
'                 paste operation.
'   • Cell P2   – Drive a "blink" warning indicator based on P2's value.
' =============================================================================
Option Explicit

' ---------------------------------------------------------------------------
' Worksheet_Change
' Fired whenever the user (or code) writes a value into one or more cells.
' ---------------------------------------------------------------------------
Private Sub Worksheet_Change(ByVal Target As Range)

    ' -----------------------------------------------------------------------
    ' Guard: nothing to do on empty selections
    ' -----------------------------------------------------------------------
    If Target Is Nothing Then Exit Sub

    ' -----------------------------------------------------------------------
    ' Suspend events and screen updates for the duration of this handler so
    ' that any writes we perform here do not re-trigger Worksheet_Change, and
    ' so the user does not see intermediate flicker.
    ' -----------------------------------------------------------------------
    Application.EnableEvents   = False
    Application.ScreenUpdating = False

    On Error GoTo Cleanup

    ' -----------------------------------------------------------------------
    ' Section 1 – Barcode normalisation and duplicate prevention (column B)
    ' -----------------------------------------------------------------------
    Dim barcodeCol As Range
    Set barcodeCol = Intersect(Target, Me.Columns("B"))

    If Not barcodeCol Is Nothing Then
        Call HandleBarcodeChanges(barcodeCol)
    End If

    ' -----------------------------------------------------------------------
    ' Section 2 – Manual overrides (columns L and/or M)
    ' -----------------------------------------------------------------------
    Dim overrideRange As Range
    Set overrideRange = Intersect(Target, Me.Range("L:M"))

    If Not overrideRange Is Nothing Then
        Call WriteOverridesToSheet(overrideRange)
    End If

    ' -----------------------------------------------------------------------
    ' Section 3 – Formula restoration after paste
    ' Re-apply any column formulas that may have been replaced by raw values
    ' during a paste.  Only rows that have data in column A are examined so
    ' we do not write into empty trailer rows.
    ' -----------------------------------------------------------------------
    If Not Intersect(Target, Me.UsedRange) Is Nothing Then
        Call RestoreColumnFormulas(Target)
    End If

Cleanup:
    Application.ScreenUpdating = True
    Application.EnableEvents   = True

    If Err.Number <> 0 Then
        MsgBox "Worksheet_Change error " & Err.Number & ": " & Err.Description, _
               vbExclamation, "Event Handler Error"
    End If

End Sub

' ---------------------------------------------------------------------------
' Worksheet_SelectionChange
' Drives the P2-based blink indicator.  Runs on every selection move so that
' the blink state stays in sync with P2 without needing a timer.
' ---------------------------------------------------------------------------
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error Resume Next
    Call UpdateBlinkWarning
    On Error GoTo 0
End Sub

' ===========================================================================
'  PRIVATE HELPERS
' ===========================================================================

' ---------------------------------------------------------------------------
' HandleBarcodeChanges
' Normalises each changed barcode cell (Trim + UCase) and rejects duplicates.
' Uses a Scripting.Dictionary built once over the existing data so duplicate
' detection is O(n) rather than O(n²).
' ---------------------------------------------------------------------------
Private Sub HandleBarcodeChanges(ByVal changedCells As Range)

    ' Build a dictionary of all existing barcodes in column B (excluding the
    ' cells being changed so we do not compare a cell against itself).
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare   ' case-insensitive uniqueness

    Dim lastRow As Long
    lastRow = Me.Cells(Me.Rows.Count, "B").End(xlUp).Row

    Dim allBarcodes As Range
    If lastRow >= 2 Then
        Set allBarcodes = Me.Range("B2:B" & lastRow)
    End If

    If Not allBarcodes Is Nothing Then
        Dim existingCell As Range
        For Each existingCell In allBarcodes
            ' Skip cells that are part of the current change (they will be
            ' evaluated below, not used as reference data).
            If Intersect(existingCell, changedCells) Is Nothing Then
                Dim existingVal As String
                existingVal = UCase(Trim(CStr(existingCell.Value)))
                If existingVal <> "" Then
                    dict(existingVal) = existingCell.Row
                End If
            End If
        Next existingCell
    End If

    ' Process each changed cell.
    Dim cell As Range
    For Each cell In changedCells
        If cell.Row < 2 Then GoTo NextCell   ' skip header row

        Dim raw As String
        raw = Trim(CStr(cell.Value))

        ' Clear-cell is always acceptable.
        If raw = "" Then GoTo NextCell

        Dim normalised As String
        normalised = UCase(raw)

        ' Check for duplicate against other rows.
        If dict.exists(normalised) Then
            MsgBox "Duplicate barcode """ & normalised & """ already exists " & _
                   "in row " & dict(normalised) & ". Entry has been removed.", _
                   vbExclamation, "Duplicate Barcode"
            cell.ClearContents
            GoTo NextCell
        End If

        ' Write the normalised value back only when it differs (avoids a
        ' redundant second Change event even with EnableEvents = False, and
        ' keeps the undo stack cleaner).
        If cell.Value <> normalised Then
            cell.Value = normalised
        End If

        ' Register in the dictionary so that later cells in the same
        ' multi-cell paste are also checked against each other.
        dict(normalised) = cell.Row

NextCell:
    Next cell

End Sub

' ---------------------------------------------------------------------------
' WriteOverridesToSheet
' Copies (barcode, override-value) pairs from columns L/M on the active sheet
' to the OVERRIDES sheet.  Writes are batched via arrays for efficiency.
' ---------------------------------------------------------------------------
Private Sub WriteOverridesToSheet(ByVal changedCells As Range)

    Const OVERRIDES_SHEET As String = "OVERRIDES"

    ' Ensure the OVERRIDES sheet exists.
    Dim ovWs As Worksheet
    On Error Resume Next
    Set ovWs = ThisWorkbook.Worksheets(OVERRIDES_SHEET)
    On Error GoTo 0

    If ovWs Is Nothing Then
        ' Create the sheet if it is absent rather than silently failing.
        Dim lastSheet As Object
        Set lastSheet = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        Set ovWs = ThisWorkbook.Worksheets.Add(After:=lastSheet)
        ovWs.Name = OVERRIDES_SHEET
        ovWs.Range("A1").Value = "Barcode"
        ovWs.Range("B1").Value = "Override"
    End If

    ' Build a dictionary of existing override rows keyed by barcode so that
    ' we can update in-place rather than appending duplicates.
    Dim ovDict As Object
    Set ovDict = CreateObject("Scripting.Dictionary")
    ovDict.CompareMode = vbTextCompare

    Dim ovLastRow As Long
    ovLastRow = ovWs.Cells(ovWs.Rows.Count, "A").End(xlUp).Row

    Dim ovRow As Long
    For ovRow = 2 To ovLastRow
        Dim ovKey As String
        ovKey = UCase(Trim(CStr(ovWs.Cells(ovRow, "A").Value)))
        If ovKey <> "" Then
            ovDict(ovKey) = ovRow
        End If
    Next ovRow

    ' Determine which data rows are affected (unique row numbers within L:M).
    Dim affectedRows As Object
    Set affectedRows = CreateObject("Scripting.Dictionary")

    Dim c As Range
    For Each c In changedCells
        If c.Row >= 2 Then
            affectedRows(c.Row) = True
        End If
    Next c

    ' Process each affected data row.
    Dim dataRow As Variant
    For Each dataRow In affectedRows.Keys
        Dim barcode As String
        barcode = UCase(Trim(CStr(Me.Cells(dataRow, "B").Value)))
        If barcode = "" Then GoTo NextRow

        Dim overrideVal As Variant
        overrideVal = Me.Cells(dataRow, "L").Value        ' column L = primary override
        Dim secondaryOverride As Variant
        secondaryOverride = Me.Cells(dataRow, "M").Value  ' column M = secondary override

        If ovDict.exists(barcode) Then
            ' Update the existing row.
            Dim updateRow As Long
            updateRow = ovDict(barcode)
            ovWs.Cells(updateRow, "B").Value = overrideVal
            ovWs.Cells(updateRow, "C").Value = secondaryOverride
        Else
            ' Append a new row.
            ovLastRow = ovLastRow + 1
            ovWs.Cells(ovLastRow, "A").Value = barcode
            ovWs.Cells(ovLastRow, "B").Value = overrideVal
            ovWs.Cells(ovLastRow, "C").Value = secondaryOverride
            ovDict(barcode) = ovLastRow
        End If

NextRow:
    Next dataRow

End Sub

' ---------------------------------------------------------------------------
' RestoreColumnFormulas
' After a paste the user may have overwritten formula cells with plain values.
' This routine re-applies the standard column formulas to any cell in a
' formula column whose formula has been lost.
'
' Extend the ColFormulas array whenever a new formula column is added.
' ---------------------------------------------------------------------------
Private Sub RestoreColumnFormulas(ByVal changedRange As Range)

    ' Define which columns carry repeating row formulas.
    ' Format: Array(column-letter, template-formula-using-ROW()-or-relative-refs)
    ' Use {ROW} as a placeholder for the actual row number.
    Dim colDefs As Variant
    colDefs = Array( _
        Array("C", "=IFERROR(VLOOKUP($B{ROW},Products!$A:$D,2,FALSE),"""")"), _
        Array("D", "=IFERROR(VLOOKUP($B{ROW},Products!$A:$D,3,FALSE),"""")"), _
        Array("E", "=IFERROR(VLOOKUP($B{ROW},Products!$A:$D,4,FALSE),"""")") _
    )

    Dim lastDataRow As Long
    lastDataRow = Me.Cells(Me.Rows.Count, "B").End(xlUp).Row

    Dim i As Integer
    For i = LBound(colDefs) To UBound(colDefs)
        Dim colLetter As String
        colLetter = colDefs(i)(0)

        Dim colRange As Range
        Set colRange = Intersect(changedRange, Me.Columns(colLetter))
        If colRange Is Nothing Then GoTo NextCol

        Dim cell As Range
        For Each cell In colRange
            If cell.Row < 2 Or cell.Row > lastDataRow Then GoTo NextCell

            ' Only restore if this cell is missing its formula.
            If Not cell.HasFormula Then
                Dim formulaTemplate As String
                formulaTemplate = colDefs(i)(1)
                Dim restoredFormula As String
                restoredFormula = Replace(formulaTemplate, "{ROW}", CStr(cell.Row))
                cell.Formula = restoredFormula
            End If

NextCell:
        Next cell
NextCol:
    Next i

End Sub

' ---------------------------------------------------------------------------
' UpdateBlinkWarning
' Reads cell P2 and toggles a visible "warning" indicator accordingly.
' The blink effect is achieved by alternating the indicator's interior colour
' each time SelectionChange fires.
' ---------------------------------------------------------------------------
Private Sub UpdateBlinkWarning()

    Const WARNING_CELL As String = "Q2"   ' cell used as the visual indicator
    Const BLINK_ON_COLOR  As Long = RGB(255, 0, 0)    ' red  – warning active
    Const BLINK_OFF_COLOR As Long = RGB(255, 255, 0)  ' yellow – alternate blink state
    Const INACTIVE_COLOR  As Long = RGB(198, 224, 180) ' green – no warning

    Dim p2Val As Variant
    p2Val = Me.Range("P2").Value

    Dim indicatorCell As Range
    Set indicatorCell = Me.Range(WARNING_CELL)

    ' P2 is truthy when non-zero numeric or a non-empty string that is not "0"
    Dim warningActive As Boolean
    If IsNumeric(p2Val) Then
        warningActive = (CDbl(p2Val) <> 0)
    Else
        warningActive = (Trim(CStr(p2Val)) <> "" And Trim(CStr(p2Val)) <> "0")
    End If

    ' Disable events before writing so we do not re-trigger Worksheet_Change.
    Application.EnableEvents = False

    If warningActive Then
        ' Toggle between the two blink colours to create visible animation.
        ' A text symbol is also toggled so the state is legible without colour.
        If indicatorCell.Interior.Color = BLINK_ON_COLOR Then
            indicatorCell.Interior.Color = BLINK_OFF_COLOR
            indicatorCell.Value = "⚠ WARNING"
        Else
            indicatorCell.Interior.Color = BLINK_ON_COLOR
            indicatorCell.Value = "⚠ WARNING"
        End If
    Else
        indicatorCell.Interior.Color = INACTIVE_COLOR
        indicatorCell.Value = ""
    End If

    Application.EnableEvents = True

End Sub
