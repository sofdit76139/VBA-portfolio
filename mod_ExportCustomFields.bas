Attribute VB_Name = "mod_ExportCustomFields"
Sub ExportCustomFields()
' Exports the custom fields list information from the Price-Desc-Cat-Prop65 sheet to a separate workbook

Application.ScreenUpdating = False

On Error Resume Next

    Dim MyWkb As String
    
    Dim MyWkSh As String

    Dim csPath As String
    
    Dim lRw As Long
    
    MyWkSh = Worksheets("Vendor Info").Range("B2").Value & " AX Custom Field Import"
    
    MyWkb = Format(Now, "YYYY-MM-DD-HHMMSS") & " " & MyWkSh & ".xlsx"
    
    csPath = Application.ActiveWorkbook.Path & "/"
    
    Worksheets("LHAWebCustomFieldValue").Copy
    
    With ActiveWorkbook
    
        Application.DisplayAlerts = False
        
        'save the workbook
        
        .SaveAs Filename:=csPath & MyWkb
        
        'rename the worksheet to Sheet1
        
        Worksheets("LHAWebCustomFieldValue").Name = "Sheet1"
        
        'remove textboxes
        
        ActiveSheet.TextBoxes.Delete
        
        'remove all formatting after converting table to a range
        
        Dim rList As Range
 
        With Worksheets("Sheet1").ListObjects("LHAWebCustomFieldValue")
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

    Worksheets("CommandCentral").Cells(13, 8).Value = Format(Now, "mm/dd/yyyy")
    Worksheets("CommandCentral").Cells(14, 8).Value = Format(Now, "hh:mm ampm")

End Sub
