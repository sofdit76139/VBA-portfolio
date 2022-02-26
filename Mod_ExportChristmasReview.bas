Attribute VB_Name = "Mod_ExportChristmasReview"
Sub MyExportChristmasReview()

' run to export Christmas review list

Dim ws As Worksheet

Dim rng As Range

Dim MyWkb As String

Dim MyWkSh As String

Dim csPath As String


Application.ScreenUpdating = False

On Error Resume Next

'check if the sheet exists, and if it does, export and delete it

For Each ws In ThisWorkbook.Worksheets
    If ws.Name Like "* Christmas Review" Then
    
        MyWkSh = ws.Name
        
        MyWkb = MyWkSh & "_" & Format(Now, "HHMMSS") & ".xlsx"

        csPath = "C:\Users\Sofie.Dittmann\OneDrive - COMPANY\Reporting\Merchandising" & "\Christmas\"

        Worksheets(MyWkSh).Copy
        With ActiveWorkbook
            .SaveAs Filename:=csPath & MyWkb
        End With
        
        Application.DisplayAlerts = False
    
        Workbooks(MyWkb).Close

        Worksheets(MyWkSh).Delete
        
    End If
    
Next ws

'add a new sheet and copy the information over

Set ws = Worksheets.Add

ActiveSheet.Name = Format(Now, "YYYY_MM_DD") & " Christmas Review"

Set ws = ActiveSheet

For Each Row In Range("ItemWebCategories[#All]").Rows

If Row.EntireRow.Hidden = False Then
 
        If rng Is Nothing Then Set rng = Row
 
        'Returns the union of two or more ranges.
        Set rng = Union(Row, rng)
    End If
 
'Continue with next row
Next Row

rng.SpecialCells(xlCellTypeVisible).Copy Destination:=ws.Range("A1")

Dim lRw As Long

lRw = ActiveSheet.Range("A:A").SpecialCells(xlCellTypeLastCell).Row - 2

Dim i As Long

With ws
    
    For i = 1 To lRw Step 1
        
'        .Cells(i + 2, 1) = .Cells(i + 1, 2)
        .Cells(i + 2, 1).NumberFormat = "@"
        
        .Cells(i + 1, 1).Activate
        .Hyperlinks.Add Anchor:=ActiveCell, Address:="https://COMPANY.com/search?w=" & ActiveCell.Text, TextToDisplay:=ActiveCell.Text ' turn all SKUs into hyperlinks

    Next i
    
    ' autofit contents
    
    For i = 1 To ActiveSheet.UsedRange.Columns.Count

        Columns(i).EntireColumn.AutoFit

    Next i
    
    For i = 3 To ActiveSheet.UsedRange.Columns.Count

        Columns(i).ColumnWidth = "15"

    Next i
    
    Columns(6).EntireColumn.AutoFit
    
    For i = 2 To lRw
    
        Rows(i).EntireRow.AutoFit
    
    Next i

End With

With ActiveWindow
    .SplitColumn = 0
    .SplitRow = 1
End With

ActiveWindow.FreezePanes = True

Application.ScreenUpdating = True

MsgBox "The export is now complete."


End Sub

