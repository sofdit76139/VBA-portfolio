Attribute VB_Name = "Mod_SlimJimReview"
Sub SlimJimExport2()

Application.ScreenUpdating = False

On Error Resume Next

    Dim MyWkb As String
    
    Dim MyWkSh As String

    Dim csPath As String
    
'    Dim lRw As Long
    
    MyWkSh = Workbooks("merchandising_reporting.xlsm").Worksheets("SlimJim").Name
    
    MyWkb = MyWkSh & "_" & Format(Now, "YYYY-MM-DD-HHMMSS") & ".xlsx"
    
    csPath = "C:\Users\Sofie.Dittmann\OneDrive - COMPANY\Reporting\Merchandising" & "\SlimJim\"
    
    Worksheets("SlimJim").Copy
    
    With ActiveWorkbook
        Application.DisplayAlerts = False
        
        .SaveAs Filename:=csPath & MyWkb
        
        ActiveWorkbook.Queries("SlimJim").Delete
        ActiveWorkbook.Queries("SalesBasic").Delete
        
        ActiveWindow.Close False
    End With

    Worksheets("RunImport").Cells(27, 6).Value = Format(Now, "mm/dd/yyyy")
    Worksheets("RunImport").Cells(27, 7).Value = Format(Now, "hh:mm ampm")

    Sheets("RunImport").Select

    Application.ScreenUpdating = True

    MsgBox "The operation is complete."

End Sub
