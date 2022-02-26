Attribute VB_Name = "Mod_KIDTop100"
Sub MyPivotCopyFormatValues3()
'select pivot table cell first

' run for the top profit in KID

Dim ws As Worksheet
Dim pt As PivotTable
Dim rngPT As Range
Dim rngPTa As Range
Dim rngCopy As Range
Dim rngCopy2 As Range
Dim lRowTop As Long
Dim lRowsPT As Long
Dim lRowPage As Long
Dim msgSpace As String

Dim MyWkb As String

Dim MyWkSh As String

Dim csPath As String

Application.ScreenUpdating = False

On Error Resume Next
Set pt = ActiveCell.PivotTable
Set rngPTa = pt.PageRange
On Error GoTo errHandler

If pt Is Nothing Then
    MsgBox "Could not copy pivot table for active cell"
    GoTo exitHandler
End If

If pt.PageFieldOrder = xlOverThenDown Then
  If pt.PageFields.Count > 1 Then
    msgSpace = "Horizontal filters with spaces." _
      & vbCrLf _
      & "Could not copy Filters formatting."
  End If
End If

'check if the sheet exists, and if it does, export and delete it

For Each ws In ThisWorkbook.Worksheets
    If ws.Name Like "* KID Top 100" Then
    
        MyWkSh = ws.Name
        
        MyWkb = MyWkSh & "_" & Format(Now, "HHMMSS") & ".xlsx"

        csPath = "C:\Users\Sofie.Dittmann\OneDrive - COMPANY\Reporting\Merchandising" & "\Top100\"

        Worksheets(MyWkSh).Copy
        With ActiveWorkbook
            .SaveAs Filename:=csPath & MyWkb
        End With
        
        Application.DisplayAlerts = False
    
        Workbooks(MyWkb).Close

        Worksheets(MyWkSh).Delete
        
    End If
    
Next ws

Set rngPT = pt.TableRange1
lRowTop = rngPT.Rows(1).Row
lRowsPT = rngPT.Rows.Count

'add a new sheet and copy the information over

Set ws = Worksheets.Add

ActiveSheet.Name = Format(Now, "YYYY_MM_DD") & " KID Top 100"

Set rngCopy = rngPT.Resize(lRowsPT - 1)
Set rngCopy2 = rngPT.Rows(lRowsPT)

rngCopy.Copy Destination:=ws.Cells(lRowTop, 1)
rngCopy2.Copy _
  Destination:=ws.Cells(lRowTop + lRowsPT - 1, 1)

If Not rngPTa Is Nothing Then
    lRowPage = rngPTa.Rows(1).Row
    rngPTa.Copy Destination:=ws.Cells(lRowPage, 1)
End If
    
ws.Columns.AutoFit
If msgSpace <> "" Then
  MsgBox msgSpace
End If

ActiveSheet.Range("A1:A2").EntireRow.Delete

Dim lRw As Long

lRw = ActiveSheet.Range("I:I").SpecialCells(xlCellTypeLastCell).Row - 2

ActiveSheet.Range("A1").EntireColumn.Insert

Range("A1").Value = "Item #"

Range("B1").Value = "Product Name"

ActiveSheet.Range("C1").EntireColumn.Insert

Range("C1").Value = "VendName"

Range("B:B").Copy
Range("A:A").PasteSpecial (xlPasteFormats)
Range("C:C").PasteSpecial (xlPasteFormats)

ActiveSheet.Range("J1").Cells.Value = "StCost"

ActiveSheet.Range("K1").Cells.Value = "BasePrice"

ActiveSheet.Range("L1").Cells.Value = "BaseMargin"

Application.CutCopyMode = False

Dim i As Long

Set ws = ActiveSheet

With ws
    
    For i = 1 To lRw Step 3
        
        .Cells(i + 2, 1) = .Cells(i + 1, 2)
        .Cells(i + 2, 1).NumberFormat = "@"
        
    Next i
    
    For i = 1 To lRw Step 3
        
        .Cells(i + 2, 3) = .Cells(i + 3, 2)
        .Cells(i + 2, 4) = .Cells(i + 3, 4)
        .Cells(i + 2, 5) = .Cells(i + 3, 5)
        .Cells(i + 2, 6) = .Cells(i + 3, 6)
        .Cells(i + 2, 7) = .Cells(i + 3, 7)
        .Cells(i + 2, 8) = .Cells(i + 3, 8)
        .Cells(i + 2, 9) = .Cells(i + 3, 9)
        
    Next i
    
    For i = lRw To 2 Step -3

        Rows(i).Delete Shift:=xlUp
        
    Next i
    
    Columns("A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete

End With

With ws.Cells(1).CurrentRegion

    .Rows(.Rows.Count).EntireRow.Delete

End With

'writes two formulas into cells and copies them down
With ws

    Range("J2:J" & lRw).Formula = "=IFNA(VLOOKUP($A2,SalesBasic,8,false)," & Chr(34) & Chr(34) & ")"
    Range("J2:J" & lRw).NumberFormat = "_($* #,##0.00_)"
    
    Range("K2:K" & lRw).Formula = "=IFNA(VLOOKUP($A2,SalesBasic,10,false)," & Chr(34) & Chr(34) & ")"
    Range("K2:K" & lRw).NumberFormat = "_($* #,##0.00_)"
    
    Range("L2:L" & lRw).Formula = "=IFNA(VLOOKUP($A2,SalesBasic,12,false)," & Chr(34) & Chr(34) & ")"
    Range("L2:L" & lRw).NumberFormat = "0.00%"

End With

ActiveSheet.Range("A:C").IndentLevel = 0

ActiveSheet.Range("M1").Cells.Value = "PickForWeb"

ActiveSheet.Range("N1").Cells.Value = "DisregardForNext"

Range("A:K").Font.Bold = False
Range("1:1").Font.Bold = True

Range("C:C").Copy
Range("M:N").PasteSpecial (xlPasteFormats)

Range("B1").Copy
Range("J1:L1").PasteSpecial (xlPasteFormats)

ActiveSheet.Range("A:C").Columns.AutoFit
ActiveSheet.Range("K:N").Columns.AutoFit

With ActiveWindow
    .SplitColumn = 0
    .SplitRow = 1
End With

ActiveWindow.FreezePanes = True

Application.CutCopyMode = False

With ws

    Range("A1:N" & lRw).Sort Key1:=Range("C1"), _
                            Order1:=xlAscending, _
                            Key2:=Range("H1"), _
                            Order1:=xlDescending, _
                            Header:=xlYes

End With

Cells(2, 10).Select

Application.ScreenUpdating = True

MsgBox "The export is now complete."

exitHandler:
    Exit Sub
errHandler:
    MsgBox "Could not copy pivot table for active cell"
    Resume exitHandler
End Sub

