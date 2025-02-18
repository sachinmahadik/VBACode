Option Explicit

Sub DispatchReportMaker()

    ' Declaring necessary variables
    Dim val As String
    Dim companyname As String
    Dim tbl As ListObject
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim dt As Date
    Dim msgtext As String

    ' Error handling and screen updating
    Application.ScreenUpdating = False
    On Error GoTo ErrHandler

    ' Assigning the workbook and worksheet
    Set wb = ActiveWorkbook
    Set ws = wb.Sheets("Data")

    ' Determine the report type based on the workbook name
    val = wb.Name
    If val = "Asian Paints - Daily Dispatch- Report.xlsx" Then
        GoTo Program
    ElseIf val = "Weekly Open order Report.xlsx" Then
        msgtext = MsgBox("You are in Open Order File, do you want to generate the report?", vbYesNo)
        If msgtext = vbYes Then
            GoTo Program3
        Else
            Exit Sub
        End If
    Else
        msgtext = MsgBox("You are not in an Asian Dispatch Report File, but still want to generate the report?", vbYesNo)
        If msgtext = vbYes Then
            GoTo Program2
        Else
            Exit Sub
        End If
    End If

Program:
    ' Assigning company name for saving the new file
    companyname = "APL - Daily Dispatch Report"

    ' Creating the new report based on the existing data
    Set tbl = ws.ListObjects("datasales")
    Call CreateDispatchReport(tbl, companyname, "D:\SMAHADIK_WD\sachin\Excel\Asian Daily Dispatch\")
    Exit Sub

Program2:
    ' Assign company name dynamically for a different report
    companyname = Left(wb.Name, (InStrRev(wb.Name, ".", -1, vbTextCompare) - 1))

    ' Create report for non-Asian Paints files
    Set tbl = ws.ListObjects("datasales")
    Call CreateDispatchReport(tbl, companyname, "D:\SMAHADIK_WD\sachin\Excel\Other Dispatch Report\")
    Exit Sub

Program3:
    ' Creating a report for Weekly Open Order file
    companyname = Left(wb.Name, (InStrRev(wb.Name, ".", -1, vbTextCompare) - 1))
    Set tbl = ws.ListObjects("OpenOrder")
    Call CreateDispatchReport(tbl, companyname, "D:\SMAHADIK_WD\sachin\Excel\Weekly Open Order\")
    Exit Sub

ErrHandler:
    MsgBox "An error occurred. Please check if the file is already open or if there are any issues.", vbOKOnly, "Error"
    Application.ScreenUpdating = True
    Exit Sub

End Sub

' --------------------------- HELPER FUNCTION ---------------------------

Sub CreateDispatchReport(tbl As ListObject, companyname As String, folderPath As String)
    Dim wbn As Workbook
    Dim dt As Date
    Dim nameofwb As String
    Dim src As Range

    ' Get current date and generate file name
    dt = Date
    nameofwb = companyname & " - " & dt

    ' Create new workbook
    Set wbn = Workbooks.Add
    wbn.Worksheets("Sheet1").Name = "Data"

    ' Copy the sales data into the new workbook
    tbl.Range.Copy
    wbn.Worksheets("Data").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats

    ' Add the data to a ListObject (table) in the new workbook
    Set src = wbn.Worksheets("Data").UsedRange
    wbn.Worksheets("Data").ListObjects.Add SourceType:=xlSrcRange, Source:=src, _
        xlListObjectHasHeaders:=xlYes, TableStyleName:="TableStyleMedium28"

    ' Save the new workbook
    wbn.SaveAs fileName:=folderPath & nameofwb
    MsgBox "Report created successfully at " & folderPath & nameofwb, vbInformation, "Report Created"

    ' Clean up
    Set wbn = Nothing
    Set src = Nothing
End Sub


' --------------------------- HELPER FUNCTION ---------------------------

Sub CreateDispatchReport(tbl As ListObject, companyname As String, folderPath As String)
    Dim wbn As Workbook
    Dim dt As Date
    Dim nameofwb As String
    Dim src As Range

    ' Get current date and generate file name
    dt = Date
    nameofwb = companyname & " - " & dt

    ' Create new workbook
    Set wbn = Workbooks.Add
    wbn.Worksheets("Sheet1").Name = "Data"

    ' Copy the sales data into the new workbook
    tbl.Range.Copy
    wbn.Worksheets("Data").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats

    ' Add the data to a ListObject (table) in the new workbook
    Set src = wbn.Worksheets("Data").UsedRange
    wbn.Worksheets("Data").ListObjects.Add SourceType:=xlSrcRange, Source:=src, _
        xlListObjectHasHeaders:=xlYes, TableStyleName:="TableStyleMedium28"

    ' Save the new workbook
    wbn.SaveAs fileName:=folderPath & nameofwb
    MsgBox "Report created successfully at " & folderPath & nameofwb, vbInformation, "Report Created"

    ' Clean up
    Set wbn = Nothing
    Set src = Nothing
End Sub