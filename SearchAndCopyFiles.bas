Option Explicit

' This macro searches for files in a source folder 
' based on a list in an Excel sheet and copies them to 
' a destination folder. this micro is created to copy the files from one folder to another folder.
' since we have centralised dump of all invoices. Then only copy the invoice which are required in list.


Sub SearchAndCopyFiles()

    ' Declare necessary variables
    Dim cell As Range
    Dim sourceFolderPath As String
    Dim destinationFolderPath As String
    Dim fileFound As Boolean
    Dim fso As Object
    Dim sourceFile As String
    Dim fileCollection As Object
    Dim file As Object
    Dim fileName As String
    Dim finalrow As Long
    Dim NewFldrName As String
    Dim CurrentDate As String
    Dim Msg As String
    Dim fldr As FileDialog

    ' Ensure the current workbook is an AR Report
    If InStr(ActiveWorkbook.Name, "AR Report") > 0 Then
        NewFldrName = Trim(Left(ActiveWorkbook.Name, InStr(ActiveWorkbook.Name, "-") - 1))
        GoTo StartProcess
    Else
        Msg = MsgBox("Please select the AR Report file, then run the script or select OK.", vbOKCancel, "ALERT")
        If Msg = vbOK Then
            GoTo StartProcess
        Else
            Exit Sub
        End If
    End If

StartProcess:
    ' Prompt user to select the source folder
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    fldr.Show
    
    ' Set current date as folder name
    CurrentDate = Format(Now(), "yyyy-mm-dd")
    
    ' Set source folder and destination folder paths
    sourceFolderPath = fldr.SelectedItems(1)
    destinationFolderPath = "C:\Users\sachin.mahadik\Desktop\Sub"

    ' Create a new folder in the "Submission" folder with the company name and current date
    If Dir(destinationFolderPath & "\" & NewFldrName & "\" & CurrentDate, vbDirectory) = "" Then
        MkDir destinationFolderPath & "\" & NewFldrName & "\" & CurrentDate
    End If
    destinationFolderPath = destinationFolderPath & "\" & NewFldrName & "\" & CurrentDate

    ' Ensure folder paths end with a backslash
    If Right(sourceFolderPath, 1) <> "\" Then sourceFolderPath = sourceFolderPath & "\"
    If Right(destinationFolderPath, 1) <> "\" Then destinationFolderPath = destinationFolderPath & "\"

    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Find the last row in the active sheet
    finalrow = ActiveSheet.UsedRange.Rows.Count

    ' Loop through each cell in column F to find file names
    For Each cell In ActiveSheet.Range("F1:F" & finalrow)
        fileName = Trim(cell.Value)
        
        ' Skip if cell is empty or invalid
        If Len(fileName) < 5 Then
            cell.Offset(0, 15).Value = "Invalid Name"
            GoTo NextCell
        End If
        
        fileFound = False
        
        ' Get file collection from source folder
        Set fileCollection = fso.GetFolder(sourceFolderPath).Files
        
        ' Search through each file in the source folder
        For Each file In fileCollection
            ' If file name contains the search term (cell value)
            If InStr(file.Name, fileName) > 0 Then
                sourceFile = sourceFolderPath & file.Name
                
                ' Copy file to destination folder
                fso.CopyFile sourceFile, destinationFolderPath & file.Name
                fileFound = True
                cell.Offset(0, 15).Value = "Copied" ' Update status
                Exit For ' Exit once the file is copied
            End If
        Next file
        
        ' If file not found, mark as "Not Found"
        If Not fileFound Then
            cell.Offset(0, 15).Value = "Not Found"
        End If
        
NextCell:
    Next cell

    ' Cleanup
    Set fso = Nothing
    Set fldr = Nothing

    ' Notify the user the process is complete
    MsgBox "File search and copy process completed.", vbInformation, "Completed"
    
End Sub

