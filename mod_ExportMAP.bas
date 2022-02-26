Attribute VB_Name = "mod_ExportMAP"
Sub ExportMAP()
' Exports the ProductInfoAX sheet to a separate workbook and deletes unneeded columns

Application.ScreenUpdating = False

On Error Resume Next

    Dim MyWkb As String
    
    Dim MyWkSh As String

    Dim csPath As String
    
'    Dim lRw As Long
    
    MyWkSh = Worksheets("Vendor Info").Range("B2").Value & " MAP Changes"
    
    MyWkb = Format(Now, "YYYY-MM-DD-HHMMSS") & " " & MyWkSh & ".xlsx"
    
    csPath = Application.ActiveWorkbook.Path & "/"
    
    Worksheets("ProductInfoAX").Copy
    
    With ActiveWorkbook
    
        Application.DisplayAlerts = False
        
        'save the workbook
        
        .SaveAs Filename:=csPath & MyWkb
        
        'remove all conditional formatting
        
        Columns("E:E").FormatConditions.Delete
        
        'remove unwanted columns
        
        Columns("F:K").Delete
        Columns("B:D").Delete
        
        'remove textboxes
        
        ActiveSheet.TextBoxes.Delete
        
        'change column headings
        
        ActiveSheet.Range("A1").Value = "ItemID"
        ActiveSheet.Range("B1").Value = "LHAMAPPrice"
        
        Dim rList As Range
 
        With Worksheets("Sheet1").ListObjects("ProductInfo")
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

    MsgBox "A MAP file was generated for batch upload."

End Sub
