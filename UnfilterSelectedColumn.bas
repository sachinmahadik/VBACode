Sub UnfilterSelectedColumn()
    '   How It Works:
    '   Select a single cell in the column you want to unfilter.
    '   Run the macro.
    '   It removes the filter from that specific column, whether it's inside a Table or a regular range.
    '   If the column isn't filtered, it shows a message saying "No filter found.

    Dim ws As Worksheet
    Dim selectedCell As Range
    Dim tbl As ListObject
    Dim filterColumn As Long

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
        
        ' Remove filter from the specific column
        If tbl.AutoFilter.FilterMode Then
            tbl.Range.AutoFilter Field:=filterColumn
        Else
            MsgBox "No active filter found in this column.", vbInformation, "No Filter"
        End If
    Else
        ' The selected cell is inside a regular range
        If ws.AutoFilterMode = False Then
            MsgBox "No active filters found.", vbInformation, "No Filter"
            Exit Sub
        End If

        ' Get the filter column index
        Dim dataRange As Range
        Set dataRange = selectedCell.CurrentRegion
        filterColumn = selectedCell.Column - dataRange.Columns(1).Column + 1

        ' Check if the column has an active filter and remove it
        If dataRange.Columns.Count >= filterColumn Then
            dataRange.AutoFilter Field:=filterColumn
        Else
            MsgBox "No filter found in this column.", vbInformation, "No Filter"
        End If
    End If
End Sub
