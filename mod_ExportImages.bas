Attribute VB_Name = "mod_ExportImages"
Sub ExportImages()
' Exports the Images sheet to a separate workbook and deletes unneeded items

Application.ScreenUpdating = False

On Error Resume Next

    Dim MyWkb As String
    
    Dim MyWkSh As String

    Dim csPath As String
    
    Dim lRw As Long
    
    MyWkSh = Worksheets("Vendor Info").Range("B2").Value & " AX Image Import"
    
    MyWkb = Format(Now, "YYYY-MM-DD-HHMMSS") & " " & MyWkSh & ".xlsx"
    
    csPath = Application.ActiveWorkbook.Path & "/"
    
    Worksheets("Images").Copy
    
    With ActiveWorkbook
    
        Application.DisplayAlerts = False
        
        'save the workbook
        
        .SaveAs Filename:=csPath & MyWkb
        
        'remove textboxes
        
        ActiveSheet.TextBoxes.Delete
        
        'rename the worksheet to Sheet1
        
        Worksheets("Images").Name = "Sheet1"
        
        'remove all formatting after converting table to a range
        
        Dim rList As Range
 
        With Worksheets("Sheet1").ListObjects("Images")
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

        ActiveWindow.Close False
    End With

    Application.ScreenUpdating = True

'    MsgBox "The import file was generated for batch upload."

    Worksheets("CommandCentral").Cells(13, 20).Value = Format(Now, "mm/dd/yyyy")
    Worksheets("CommandCentral").Cells(14, 20).Value = Format(Now, "hh:mm ampm")

End Sub
