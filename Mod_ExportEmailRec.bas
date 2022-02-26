Attribute VB_Name = "Mod_ExportEmailRec"
Sub NextMonthEmail()

Application.ScreenUpdating = False

On Error Resume Next

    Dim MyWkb As String
    
    Dim MyWkSh As String

    Dim csPath As String
    
    csPath = "C:\Users\Sofie.Dittmann\OneDrive - Lehman's\Reporting\Merchandising" & "\Emails\"
    
    MyWkSh = Workbooks("merchandising_reporting.xlsm").Worksheets("NextMonthEmail").Name
    
    MyWkb = MyWkSh & "_" & Format(Now, "YYYY-MM-DD-HHMMSS") & ".xlsx"
    
    Workbooks("merchandising_reporting.xlsm").Worksheets("NextMonthEmail").Copy
    With ActiveWorkbook
        .SaveAs Filename:=csPath & MyWkb
    End With
    
    Application.DisplayAlerts = False
    
    Workbooks(MyWkb).Close
    
    MyWkSh = Workbooks("merchandising_reporting.xlsm").Worksheets("NextMonthEmail").Name
    
    MyWkb = MyWkSh & "_" & Format(Now, "YYYY-MM-DD-HHMMSS") & ".xlsx"
    
    Workbooks("merchandising_reporting.xlsm").Activate
    
    Worksheets("SKUs for Emails").Copy
    With ActiveWorkbook
        .SaveAs Filename:=csPath & MyWkb
    End With
    
    Application.DisplayAlerts = False
    
    Workbooks(MyWkb).Close
    
    'if there are any hidden columns, unhide them:
    
    On Error Resume Next
    
    Sheet25.ShowAllData
    
    On Error GoTo 0
    
    'clear the last two columns in this sheet:
    
    Windows("merchandising_reporting.xlsm").Activate
    
    Worksheets("NextMonthEmail").Activate
    
    Dim lRw As Long
    
    lRw = Range("W:X").SpecialCells(xlCellTypeLastCell).Row
    
    Workbooks("merchandising_reporting.xlsm").Worksheets("NextMonthEmail").Range("W2:X" & lRw).Clear
    
    'clear all contents in columns F - Q in this sheet
    
    Worksheets("SKUs for Emails").Activate
    
    lRw = Range("F:Q").SpecialCells(xlCellTypeLastCell).Row
    
    Workbooks("merchandising_reporting.xlsm").Worksheets("SKUs for Emails").Range("F2:Q" & lRw).ClearContents

    Workbooks.Open (ThisWorkbook.Path & "\Lst Yr Month Lst Yr Qtr ISDT Report.xlsx")
    
    'clean up the original report and refresh the query

    Windows("Lst Yr Month Lst Yr Qtr ISDT Report.xlsx").Activate

    Sheets(1).Name = "NextMonth"

    ActiveWindow.SmallScroll Down:=-12
    ActiveWindow.FreezePanes = False
    ActiveWindow.SmallScroll Down:=-15

    Rows("1:4").Delete Shift:=xlUp

    Application.DisplayAlerts = False

    Windows("Lst Yr Month Lst Yr Qtr ISDT Report.xlsx").Activate

    ActiveWindow.Close False
    
    Worksheets("SKUs for Emails").Activate
    
    Dim MyDateYear As String
    Dim MyDateMonth As String
    
    MyDateYear = Format(Now, "YYYY")
    MyDateMonth = Format(Now, "MM")
    
    Cells(2, 6).FormulaArray = "=SEQUENCE(32,1, DATE(" & MyDateYear & "," & MyDateMonth & ",1),1)"
    Range("F3:F" & lRw).FillDown

    Worksheets("RunImport").Cells(22, 6).Value = Format(Now, "mm/dd/yyyy")
    Worksheets("RunImport").Cells(22, 7).Value = Format(Now, "hh:mm ampm")

    ThisWorkbook.RefreshAll
    
    Range("A:Y" & lRw).Sort Key1:=Range("T2:T" & lRw), _
    Order1:=xlAscending, Header:=xlNo
    
    'hide columns
    
    Worksheets("NextMonthEmail").Range("C:C").EntireColumn.Hidden = True
    Worksheets("NextMonthEmail").Range("G:G").EntireColumn.Hidden = True
    Worksheets("NextMonthEmail").Range("J:J").EntireColumn.Hidden = True

    Sheets("RunImport").Select

    Application.ScreenUpdating = True

    MsgBox "The operation is complete."

End Sub
