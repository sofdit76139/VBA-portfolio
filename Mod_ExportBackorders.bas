Attribute VB_Name = "Mod_ExportBackorders"
Sub ExportBackorders()

Application.ScreenUpdating = False

On Error Resume Next

    Dim MyWkb As String
    
    Dim MyWkSh As String

    Dim csPath As String
    
    Dim lRw As Long
    
    MyWkSh = Workbooks("merchandising_reporting.xlsm").Worksheets("NewArrivalBackorders").Name
    
    MyWkb = MyWkSh & "_" & Format(Now, "YYYY-MM-DD-HHMMSS") & ".xlsx"
    
    csPath = "C:\Users\Sofie.Dittmann\OneDrive - COMPANY\Reporting\Merchandising" & "\BackOrders\"
    
    Worksheets("NewArrivalBackorders").Copy
    With ActiveWorkbook
        Application.DisplayAlerts = False
        
        .SaveAs Filename:=csPath & MyWkb
        
'        ActiveWorkbook.Queries("NewArrivalBackorders").Delete
'        ActiveWorkbook.Queries("DirectSalesLessMktPlaces").Delete
        
        ActiveWindow.Close False
    End With
    
    Worksheets("RunImport").Cells(23, 6).Value = Format(Now, "mm/dd/yyyy")
    Worksheets("RunImport").Cells(23, 7).Value = Format(Now, "hh:mm ampm")

    Sheets("RunImport").Select

    Application.ScreenUpdating = True

    MsgBox "The operation is complete."

End Sub


