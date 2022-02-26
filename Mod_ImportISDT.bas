Attribute VB_Name = "Mod_ImportISDT"
Sub ImportData()
'
' ImportData Macro to import the ISDT YTD table and divide it up
' Copy the latest version of this report into the same directory as this workbook before running the macro
'

'01/26/2022

Application.ScreenUpdating = False

On Error Resume Next

    Workbooks.Open (ThisWorkbook.Path & "\ItemSalesDataTable FullYear.xlsb")

    Windows("merchandising_reporting.xlsb").Activate
    
' Start by clearing all data from the sheets in our workbook, with the exception of the first one, "RunImport"
    
    Worksheets("Sales Basic").Cells.Clear
    Worksheets("Direct Sales Less Mkt Places").Cells.Clear
    Worksheets("Kidron Sales").Cells.Clear
    Worksheets("Direct Sales Less Mkt Places").Cells.Clear
    Worksheets("Market Place Sales").Cells.Clear
'    Sheets.Add(After:=Sheets("Direct Sales Less Mkt Places")).Name = "Market Place Sales"
'    Sheets.Add(After:=Sheets("Market Place Sales")).Name = "Direct Sales"

' Change to the other workbook, unfreeze the top rows and delete them, then refreeze the top row

    Windows("ItemSalesDataTable FullYear.xlsb").Activate

    ActiveWindow.SmallScroll Down:=-12
    ActiveWindow.FreezePanes = False
    ActiveWindow.SmallScroll Down:=-15

    Rows("1:4").Delete Shift:=xlUp
    
'   Copy the basic data into the "SalesBasic" sheet, reformat the table
    
    Columns("A:BF").Copy
    
    Windows("merchandising_reporting.xlsb").Activate
    Sheets("Sales Basic").Select
    Range("A1").Select
    ActiveSheet.Paste
    
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A:$BF"), , xlYes).Name = _
        "SalesBasic"
    Columns("A:BF").Select
    ActiveSheet.ListObjects("SalesBasic").TableStyle = "TableStyleMedium15"

    Windows("ItemSalesDataTable FullYear.xlsb").Activate
    
    Columns("A:A").Copy

    Windows("merchandising_reporting.xlsb").Activate
    
    Sheets("Direct Sales Less Mkt Places").Activate
    Range("A1").Select
    ActiveSheet.Paste

    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A:$A"), , xlYes).Name = _
        "DirectSalesLessMktPlaces"
    Columns("A:A").Select
    ActiveSheet.ListObjects("DirectSalesLessMktPlaces").TableStyle = "TableStyleMedium15"
    
    Sheets("Market Place Sales").Activate
    Range("A1").Select
    ActiveSheet.Paste

    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A:$A"), , xlYes).Name = _
        "MarketPlaceSales"
    Columns("A:A").Select
    ActiveSheet.ListObjects("MarketPlaceSales").TableStyle = "TableStyleMedium15"

    Sheets("Direct Sales").Activate
    Range("A1").Select
    ActiveSheet.Paste

    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A:$A"), , xlYes).Name = _
        "DirectSales"
    Columns("A:A").Select
    ActiveSheet.ListObjects("DirectSales").TableStyle = "TableStyleMedium15"

    Sheets("Kidron Sales").Activate
    Range("A1").Select
    ActiveSheet.Paste

    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A:$A"), , xlYes).Name = _
        "KidronSales"
    Columns("A:A").Select
    ActiveSheet.ListObjects("KidronSales").TableStyle = "TableStyleMedium15"
    
    Windows("ItemSalesDataTable FullYear.xlsb").Activate
    
    Columns("BG:BU").Copy
    
    Windows("merchandising_reporting.xlsb").Activate
    Sheets("Direct Sales Less Mkt Places").Select
    Range("B1").Select
    'Range("B1").NumberFormat = "General"
    ActiveSheet.Paste

    'resize worksheet
    With Worksheets("Direct Sales Less Mkt Places").ListObjects("DirectSalesLessMktPlaces")
        .Resize .Range(1, 1).CurrentRegion
    End With
    
    Windows("ItemSalesDataTable FullYear.xlsb").Activate

    Columns("BV:CJ").Copy
    
    Windows("merchandising_reporting.xlsb").Activate
    Sheets("Market Place Sales").Select
    Range("B1").Select
    'Range("B1").NumberFormat = "General"
    ActiveSheet.Paste
 
    'resize worksheet
    With Worksheets("Market Place Sales").ListObjects("MarketPlaceSales")
        .Resize .Range(1, 1).CurrentRegion
    End With
    
    Windows("ItemSalesDataTable FullYear.xlsb").Activate
    
    Columns("CK:CY").Copy
    
    Windows("merchandising_reporting.xlsb").Activate
    Sheets("Direct Sales").Select
    Range("B1").Select
    'Range("B1").NumberFormat = "General"
    ActiveSheet.Paste
    
    'resize worksheet
    With Worksheets("Direct Sales").ListObjects("DirectSales")
        .Resize .Range(1, 1).CurrentRegion
    End With
    
    Windows("ItemSalesDataTable FullYear.xlsb").Activate
    
    Columns("CZ:DN").Copy
    
    Windows("merchandising_reporting.xlsb").Activate
    Sheets("Kidron Sales").Select
    Range("B1").Select
    'Range("B1").NumberFormat = "General"
    ActiveSheet.Paste
    
    'resize worksheet
    With Worksheets("Kidron Sales").ListObjects("KidronSales")
        .Resize .Range(1, 1).CurrentRegion
    End With
    
    Windows("ItemSalesDataTable FullYear.xlsb").Activate
    
    Columns("DQ:DV").Copy

    Windows("merchandising_reporting.xlsb").Activate

    Sheets("Sales Basic").Select

    Range("BG1").Select
   
    ActiveSheet.Paste
    
    Windows("ItemSalesDataTable FullYear.xlsb").Activate
    
    Columns("EJ:EL").Copy

    Windows("merchandising_reporting.xlsb").Activate

    Sheets("Sales Basic").Select

    Range("BM1").Select
    
    ActiveSheet.Paste
'
'    Columns("G:I").Delete
'    Columns("L:L").Delete
'    Columns("V:V").Delete
'    Columns("AC:AC").Delete
'    Columns("AH:AH").Delete
'    Columns("AS:AT").Delete
'
'    'resize worksheet

    With Worksheets("Sales Basic").ListObjects("SalesBasic")
        .Resize .Range(1, 1).CurrentRegion
    End With

    'refresh all data connections in the workbook
'    ThisWorkbook.RefreshAll
    
    Application.DisplayAlerts = False
    
    Windows("ItemSalesDataTable FullYear.xlsb").Activate
    
    ActiveWindow.Close False
    
    Worksheets("RunImport").Cells(2, 6).Value = Format(Now, "mm/dd/yyyy")
    Worksheets("RunImport").Cells(2, 7).Value = Format(Now, "hh:mm ampm")
    
    Sheets("RunImport").Select
    
    Application.ScreenUpdating = True
    
    MsgBox "The import is now complete."
    
End Sub

