Option Explicit

' This module cleans payment data, creates a new AR report, and saves it in a specific location.
' It also generates a Pivot Table based on the required parameters.

Sub CleanPaymentData()
    
    ' Declare variables
    Dim wb As Workbook, wbn As Workbook
    Dim ws As Worksheet, sht As Worksheet
    Dim pvtCache As PivotCache
    Dim pvt As PivotTable
    Dim rg As Range
    
    Dim StartPvt As String, SrcData As String
    Dim dt As Date
    Dim nameofwb As String
    Dim msgResponse As VbMsgBoxResult
    Dim companyname As String

    ' Disable screen updating for performance
    Application.ScreenUpdating = False
    On Error GoTo ErrHandler ' Enable error handling

    ' Assign active workbook and worksheet
    Set wb = ActiveWorkbook
    Set ws = wb.ActiveSheet

    ' Ask user to confirm selection of payment data
    msgResponse = MsgBox("Have you selected the payment data to clean?", vbYesNo + vbQuestion, "Confirmation")
    If msgResponse = vbNo Then Exit Sub ' Exit if user selects "No"

    ' === Renaming Columns for Clarity ===
    ws.Range("R1").Value = "Amount"
    ws.Range("X1").Value = "Recd Date"
    ws.Range("F1").Value = "Invoice NO."
    ws.Range("G1").Value = "SAP Ref NO."
    ws.Cells(1, 22).Value = "Inv.Category"

    ' Assign company name from column C (row 2) for naming the report file
    companyname = ws.Cells(2, 3).Text

    ' === Format Columns ===
    ws.Range("R:R").NumberFormat = "#,##0" ' Format Amount column as number with thousand separator

    ' === Delete Unwanted Columns ===
    ws.Columns(Array(5, 13, 14, 15, 16, 17, 19, 20)).EntireColumn.Delete

    ' === Select & Copy Used Range ===
    Set rg = ws.UsedRange
    rg.Copy

    ' === Create New Workbook & Save with Company Name ===
    dt = Date
    nameofwb = companyname & " - AR Report " & Format(dt, "YYYY-MM-DD")

    ' Create new workbook
    Set wbn = Workbooks.Add
    wbn.SaveAs fileName:="D:\SMAHADIK_WD\sachin\Excel\AR\" & nameofwb

    ' Rename default sheet and paste cleaned data
    With wbn.Sheets(1)
        .Name = "Data"
        .Range("A1").PasteSpecial Paste:=xlPasteAll
    End With

    ' === Create Pivot Table ===
    SrcData = "'" & wbn.Sheets("Data").Name & "'!" & wbn.Sheets("Data").UsedRange.Address(ReferenceStyle:=xlR1C1)

    ' Add new worksheet for Pivot Table
    Set sht = wbn.Sheets.Add
    sht.Name = "Pivot"

    ' Define Pivot Table starting location
    StartPvt = sht.Name & "!" & sht.Range("A3").Address(ReferenceStyle:=xlR1C1)

    ' Create Pivot Cache
    Set pvtCache = wbn.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)

    ' Create Pivot Table
    Set pvt = pvtCache.CreatePivotTable(TableDestination:=StartPvt, TableName:="PivotTable1")

    ' === Configure Pivot Table Fields ===
    With pvt
        .PivotFields("Recd Date").Orientation = xlPageField
        .PivotFields("Recd Date").CurrentPage = ""
        
        .PivotFields("Agewise").Orientation = xlColumnField
        .PivotFields("Agewise").Caption = "Aging"
        
        .PivotFields("Customer").Orientation = xlRowField
        .PivotFields("Inv.Category").Orientation = xlRowField
        .PivotFields("Inv.Category").Caption = "Inv.Category"

        ' Change Compact Layout Headers
        .CompactLayoutColumnHeader = "Aging"
        .CompactLayoutRowHeader = "Company/Inv.Category"
        
        ' Add Sum of Amount
        .AddDataField .PivotFields("Amount"), "Outstanding Payments", xlSum

        ' Format Data Fields
        .DataBodyRange.NumberFormat = "#,##0"

        ' Enable auto-calculation
        .ManualUpdate = False
    End With

    ' Restore screen updating
    Application.ScreenUpdating = True

    ' Success Message
    MsgBox "AR Report has been successfully generated and saved!", vbInformation, "Success"

    Exit Sub

ErrHandler:
    ' Error Handling
    MsgBox "An error occurred: " & Err.Description, vbExclamation, "Error"
    Err.Clear
    Application.ScreenUpdating = True

End Sub
