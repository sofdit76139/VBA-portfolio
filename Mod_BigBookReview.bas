Attribute VB_Name = "Mod_BigBookReview"
Sub BigBookReview2()

Application.ScreenUpdating = False

On Error Resume Next

    Dim MyWkb As String
    
    Dim MyWkSh As String

    Dim csPath As String
    
    Dim lRw As Long
    
    Dim ws As Worksheet
    
    MyWkSh = Workbooks("merchandising_reporting.xlsm").Worksheets("BigBookReview").Name
    
    MyWkb = MyWkSh & "_" & Format(Now, "YYYY-MM-DD-HHMMSS") & ".xlsx"
    
    csPath = "C:\Users\Sofie.Dittmann\OneDrive - COMPANY\Reporting\Merchandising" & "\BigBook\"
    
    Worksheets("BigBookReview").Copy
    
    With ActiveWorkbook
        Application.DisplayAlerts = False
        
        .SaveAs Filename:=csPath & MyWkb
        
        ActiveWindow.Close False
    End With
    
    Worksheets("RunImport").Cells(31, 6).Value = Format(Now, "mm/dd/yyyy")
    Worksheets("RunImport").Cells(31, 7).Value = Format(Now, "hh:mm ampm")

    Sheets("RunImport").Select

    Application.ScreenUpdating = True

    MsgBox "The operation is complete."

End Sub

