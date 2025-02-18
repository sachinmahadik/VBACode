Sub FilterBySelectedValue()
    '   Select a single cell inside either a Table or a regular dataset.
    '   Run the macro.
    '   If the selected cell is inside a Table, it filters that column.
    '   If it's in a regular range, it applies AutoFilter accordingly.
    '   This micro should be put in shortcut and used in autofilter huge data.

    Dim ws As Worksheet
    Dim selectedCell As Range
    Dim dataRange As Range
    Dim filterColumn As Long
    Dim tbl As ListObject

    ' Set the active sheet and selected cell
    Set ws = ActiveSheet
    Set selectedCell = Selection

    ' Ensure a single cell is selected
    If selectedCell.Cells.Count > 1 Then
        MsgBox "Please select a single cell.", vbExclamation, "Invalid Selection"
        Exit Sub
    End If

    ' Check if the selected cell is inside a Table
    On Error Resume Next
    Set tbl = selectedCell.ListObject
    On Error GoTo 0

    If Not tbl Is Nothing Then
        ' The selected cell is inside a Table
        filterColumn = selectedCell.Column - tbl.Range.Columns(1).Column + 1
        tbl.Range.AutoFilter Field:=filterColumn, Criteria1:=selectedCell.Value
    Else
        ' The selected cell is inside a regular range
        On Error Resume Next
        Set dataRange = selectedCell.CurrentRegion
        On Error GoTo 0

        If dataRange Is Nothing Then
            MsgBox "No data found around the selected cell.", vbExclamation, "Error"
            Exit Sub
        End If

        ' Determine the column number of the selected cell within the range
        filterColumn = selectedCell.Column - dataRange.Columns(1).Column + 1

        ' Apply AutoFilter based on selected cell value
        If ws.AutoFilterMode = False Then
            dataRange.AutoFilter
        End If

        dataRange.AutoFilter Field:=filterColumn, Criteria1:=selectedCell.Value
    End If
End Sub
