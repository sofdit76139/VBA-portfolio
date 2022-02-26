Attribute VB_Name = "mod_ExportCategories"
Sub ExportCategories()
' Exports the category information from the Price-Desc-Cat-Prop65 sheet to a separate text file

Application.ScreenUpdating = False

On Error Resume Next

    Dim MyWkb As String
    
    Dim MyWkSh As String

    Dim csPath As String
    
    Dim lRw As Long
    
    MyWkSh = Worksheets("Vendor Info").Range("B2").Value & " Category ID CV3 Import"
    
    MyWkb = Format(Now, "YYYY-MM-DD-HHMMSS") & " " & MyWkSh & ".txt"
    
    csPath = Application.ActiveWorkbook.Path & "/"
    
    Worksheets("Price-Desc-Cat-Prop65").Copy
    
    With ActiveWorkbook
    
        Application.DisplayAlerts = False
        
        'duplicate the SKU2 column as values

        lRw = ActiveSheet.Range("A:A").SpecialCells(xlCellTypeLastCell).Row
        
        Range("N1").EntireColumn.Insert
        
        Range("O1:O" & lRw).Copy
        
        Range("N1").PasteSpecial Paste:=xlPasteValues
        
        'remove unwanted columns
        
        Columns("R:U").Delete
        Columns("O:O").Delete
        Columns("A:M").Delete
        
        'change column headings
        
        ActiveSheet.Range("A1").Value = "SKU"
        
        'remove all formatting after converting table to a range
        
        Dim rList As Range
 
        With ActiveSheet.ListObjects("Price_Desc_Cat_Prop65")
            Set rList = .Range
            .Unlist                           ' convert the table back to a range
        End With
        
        With rList
            .Interior.ColorIndex = xlColorIndexNone
            .Font.ColorIndex = xlColorIndexAutomatic
            .Borders.LineStyle = xlLineStyleNone
            .ClearFormats
        End With
        
        'remove any data connections from the original workbook
        
        For i = 1 To ActiveWorkbook.Connections.Count
            If ActiveWorkbook.Connections.Count = 0 Then
            
                Exit For
            
            Else
                ActiveWorkbook.Connections.Item(i).Delete
            End If
            
            i = i - 1
            
        Next i
        
        'save the workbook
        
        .SaveAs Filename:=csPath & MyWkb, FileFormat:=xlTextWindows

        ActiveWindow.Close False
    End With

    Application.ScreenUpdating = True

'    MsgBox "The import file was generated for batch upload."

    Worksheets("CommandCentral").Cells(13, 11).Value = Format(Now, "mm/dd/yyyy")
    Worksheets("CommandCentral").Cells(14, 11).Value = Format(Now, "hh:mm ampm")

End Sub
