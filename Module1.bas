Attribute VB_Name = "Module1"
' =============================================================================
' Module1 – General macros
'
' SortData
'   Sorts the data table on the active sheet by the barcode column (B),
'   ascending.  Wraps the sort in performance guards (events, screen updates,
'   and calculation mode) and restores them safely even when an error occurs.
' =============================================================================
Option Explicit

' ---------------------------------------------------------------------------
' SortData
' ---------------------------------------------------------------------------
Public Sub SortData()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' ------------------------------------------------------------------
    ' Performance guards – disable until the sort is complete.
    ' ------------------------------------------------------------------
    Dim prevCalc As XlCalculation
    prevCalc = Application.Calculation

    Application.EnableEvents    = False
    Application.ScreenUpdating  = False
    Application.Calculation     = xlCalculationManual

    On Error GoTo Cleanup

    ' ------------------------------------------------------------------
    ' Determine the extent of the data (column A is the anchor column).
    ' ------------------------------------------------------------------
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    If lastRow < 2 Then
        ' Nothing to sort (header only or empty sheet).
        GoTo Cleanup
    End If

    Dim dataRange As Range
    Set dataRange = ws.Range("A1").CurrentRegion

    ' ------------------------------------------------------------------
    ' Sort ascending by barcode (column B), treating row 1 as the header.
    ' ------------------------------------------------------------------
    With dataRange.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range("B2:B" & lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal
        .Header      = xlYes
        .MatchCase   = False
        .Orientation = xlTopToBottom
        .Apply
    End With

Cleanup:
    ' ------------------------------------------------------------------
    ' Always restore application state, even when an error occurred.
    ' ------------------------------------------------------------------
    Application.Calculation    = prevCalc
    Application.ScreenUpdating = True
    Application.EnableEvents   = True

    If Err.Number <> 0 Then
        MsgBox "SortData error " & Err.Number & ": " & Err.Description, _
               vbExclamation, "Sort Error"
    End If

End Sub
