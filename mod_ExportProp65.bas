Attribute VB_Name = "mod_ExportProp65"
Sub ExportProp65()
' Exports the Prop65 information from the Price-Desc-Cat-Prop65 sheet to a separate text file

Application.ScreenUpdating = False

On Error Resume Next

    Dim MyWkb As String
    
    Dim MyWkSh As String

    Dim csPath As String
    
    Dim lRw As Long
    
    MyWkSh = Worksheets("Vendor Info").Range("B2").Value & " Prop65 CV3 Import"
    
    MyWkb = Format(Now, "YYYY-MM-DD-HHMMSS") & " " & MyWkSh & ".txt"
    
    csPath = Application.ActiveWorkbook.Path & "/"
    
    Worksheets("Price-Desc-Cat-Prop65").Copy
    
    With ActiveWorkbook
    
        Application.DisplayAlerts = False
        
        'duplicate the SKU2 column as values

        lRw = ActiveSheet.Range("A:A").SpecialCells(xlCellTypeLastCell).Row
        
        Range("Q1").EntireColumn.Insert
        
        Range("R1:R" & lRw).Copy
        
        Range("Q1").PasteSpecial Paste:=xlPasteValues
        
        'remove unwanted columns
        
        Columns("T:U").Delete
        Columns("R:R").Delete
        Columns("A:P").Delete
        
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

    Worksheets("CommandCentral").Cells(13, 14).Value = Format(Now, "mm/dd/yyyy")
    Worksheets("CommandCentral").Cells(14, 14).Value = Format(Now, "hh:mm ampm")

End Sub

