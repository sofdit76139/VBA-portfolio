Attribute VB_Name = "Mod_DeleteUnneededColumns"
Sub DeleteMyColumns()

'    Application.ScreenUpdating = False
'
'    On Error Resume Next

    Windows("merchandising_reporting.xlsm").Activate
    
    Sheets("Sales Basic").Select

    Columns("BD:BE").Delete
    Columns("AO:AO").Delete
    Columns("P:P").Delete
    Columns("H:J").Delete

    'resize worksheet
    With Worksheets("Sales Basic").ListObjects("SalesBasic")
        .Resize .Range(1, 1).CurrentRegion
    End With

    Dim sht As Worksheet
    Dim fnd As Variant
    Dim rplc As Variant
    
    fnd = "1/1/1900"
    rplc = ""
    
    'Store a specfic sheet to a variable
    Set sht = Sheets("Sales Basic")
    
    'Perform the Find/Replace All
    sht.Cells.Replace What:=fnd, Replacement:=rplc, _
    LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False

    Application.DisplayAlerts = False

    Worksheets("Market Place Sales").Delete
    Worksheets("Direct Sales").Delete
    
    'export SalesBasic, Kidron Sales, and DIR Sales Less Mkt Places to CSV
    
    Dim MyWkb As String
    
    Dim MyWkSh As String

    Dim csPath As String
    
    Dim ws As Worksheet
    
For Each ws In ThisWorkbook.Worksheets
    If ws.Name = "Sales Basic" Then
    
        MyWkSh = ws.Name
        
        MyWkb = MyWkSh & ".txt"

        csPath = "C:\Users\Sofie.Dittmann\OneDrive - Lehman's\Reporting\Merchandising\"

        Worksheets(MyWkSh).Copy
        With ActiveWorkbook
            .SaveAs Filename:=csPath & MyWkb, FileFormat:=xlText, CreateBackup:=False
        End With
        
        Application.DisplayAlerts = False
    
        Workbooks(MyWkb).Close
        
    ElseIf ws.Name = "Kidron Sales" Then
    
        MyWkSh = ws.Name
        
        MyWkb = MyWkSh & ".txt"

        csPath = "C:\Users\Sofie.Dittmann\OneDrive - Lehman's\Reporting\Merchandising\"

        Worksheets(MyWkSh).Copy
        With ActiveWorkbook
            .SaveAs Filename:=csPath & MyWkb, FileFormat:=xlText, CreateBackup:=False
        End With
        
        Application.DisplayAlerts = False
    
        Workbooks(MyWkb).Close
        
    ElseIf ws.Name = "Direct Sales Less Mkt Places" Then
    
        MyWkSh = ws.Name
        
        MyWkb = MyWkSh & ".txt"

        csPath = "C:\Users\Sofie.Dittmann\OneDrive - Lehman's\Reporting\Merchandising\"

        Worksheets(MyWkSh).Copy
        With ActiveWorkbook
            .SaveAs Filename:=csPath & MyWkb, FileFormat:=xlText, CreateBackup:=False
        End With
        
        Application.DisplayAlerts = False
    
        Workbooks(MyWkb).Close
        
        Exit For
        
    End If
    
    
Next ws
    
    Application.DisplayAlerts = True

    Workbooks("merchandising_reporting.xlsm").Worksheets("RunImport").Select
         
    Worksheets("RunImport").Cells(14, 6).Value = Format(Now, "mm/dd/yyyy")
    Worksheets("RunImport").Cells(14, 7).Value = Format(Now, "hh:mm ampm")
    
    Application.ScreenUpdating = True
    
    MsgBox "The column deletion is now complete."

End Sub
