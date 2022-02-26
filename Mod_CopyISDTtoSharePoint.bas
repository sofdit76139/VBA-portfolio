Attribute VB_Name = "Mod_CopyISDTtoSharePoint"
Sub MySaveToSharePoint()
'
' SaveToSharePoint Macro
'
'

    Dim MyWkb As String

    Const csPath As String = "C:\Users\Sofie.Dittmann\OneDrive - COMPANY/Merchandising Documents\Reports\WeeklyISDT\"
    
    'Dim arr1()
     
Application.ScreenUpdating = False

On Error Resume Next

    Workbooks.Open (ThisWorkbook.Path & "\ISDT_divided.xlsx")

    Windows("ISDT_divided.xlsx").Activate
    
    Workbooks("ISDT_divided.xlsx").Worksheets("Sales Basic").Cells.Clear
    Workbooks("ISDT_divided.xlsx").Worksheets("Direct Sales Less Mkt Places").Cells.Clear
    Workbooks("ISDT_divided.xlsx").Worksheets("Market Place Sales").Cells.Clear
    Workbooks("ISDT_divided.xlsx").Worksheets("Direct Sales").Cells.Clear
    Workbooks("ISDT_divided.xlsx").Worksheets("Kidron Sales").Cells.Clear

    Windows("merchandising_reporting.xlsm").Activate
    
'   Copy the basic data into the "SalesBasic" sheet from one file to the other, reformat the table
    
    Sheets("Sales Basic").Select

    Columns("A:BN").Copy
    
    Windows("ISDT_divided.xlsx").Activate
    Sheets("Sales Basic").Select
    Range("A1").Select
    ActiveSheet.Paste
    
'    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A:$BB"), , xlYes).Name = _
'        "SalesBasic"
'    Columns("A:BB").Select
'    ActiveSheet.ListObjects("SalesBasic").TableStyle = "TableStyleMedium15"

    Windows("merchandising_reporting.xlsm").Activate
    
    Sheets("Sales Basic").Select
    
    Columns("A:A").Copy

    Windows("ISDT_divided.xlsx").Activate
    
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
    
    Windows("merchandising_reporting.xlsm").Activate
    
    Sheets("Direct Sales Less Mkt Places").Select
    
    Columns("B:P").Copy
    
    Windows("ISDT_divided.xlsx").Activate
    Sheets("Direct Sales Less Mkt Places").Select
    Range("B1").Select
    'Range("B1").NumberFormat = "General"
    ActiveSheet.Paste

    'resize worksheet
    With Worksheets("Direct Sales Less Mkt Places").ListObjects("DirectSalesLessMktPlaces")
        .Resize .Range(1, 1).CurrentRegion
    End With
    
    Windows("merchandising_reporting.xlsm").Activate
    
    Sheets("Market Place Sales").Select
    
    Columns("B:P").Copy
    
    Windows("ISDT_divided.xlsx").Activate
    Sheets("Market Place Sales").Select
    Range("B1").Select
    'Range("B1").NumberFormat = "General"
    ActiveSheet.Paste
 
    'resize worksheet
    With Worksheets("Market Place Sales").ListObjects("MarketPlaceSales")
        .Resize .Range(1, 1).CurrentRegion
    End With
    
    Windows("merchandising_reporting.xlsm").Activate
    
    Sheets("Direct Sales").Select
    
    Columns("B:P").Copy
    
    Windows("ISDT_divided.xlsx").Activate
    Sheets("Direct Sales").Select
    Range("B1").Select
    'Range("B1").NumberFormat = "General"
    ActiveSheet.Paste
    
    'resize worksheet
    With Worksheets("Direct Sales").ListObjects("DirectSales")
        .Resize .Range(1, 1).CurrentRegion
    End With
    
    Windows("merchandising_reporting.xlsm").Activate
    
    Sheets("Kidron Sales").Select
    
    Columns("B:P").Copy
    
    Windows("ISDT_divided.xlsx").Activate
    Sheets("Kidron Sales").Select
    Range("B1").Select
    'Range("B1").NumberFormat = "General"
    ActiveSheet.Paste
    
    'resize worksheet
    With Worksheets("Kidron Sales").ListObjects("KidronSales")
        .Resize .Range(1, 1).CurrentRegion
    End With
         
    'refresh all data connections in the ISDT_divided workbook
    Workbooks("ISDT_divided.xlsx").RefreshAll
         
    Application.DisplayAlerts = False
         
' Save it with the new name and specific path

    MyWkb = Workbooks("ISDT_divided.xlsx").Name

    Workbooks("ISDT_divided.xlsx").SaveCopyAs Filename:=csPath & MyWkb
         
    Windows("ISDT_divided.xlsx").Activate

    ActiveWindow.Close False
         
    Worksheets("RunImport").Cells(10, 6).Value = Format(Now, "mm/dd/yyyy")
    Worksheets("RunImport").Cells(10, 7).Value = Format(Now, "hh:mm ampm")
    
    Sheets("RunImport").Select
    
    Application.ScreenUpdating = True
    
    MsgBox "The export is now complete."

End Sub

