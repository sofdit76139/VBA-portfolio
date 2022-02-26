Attribute VB_Name = "mod_ExportParentSKUs"
Sub ExportParentSKUs()
' Exports the parent SKU information from the ParentSKUs sheet to a separate workbook

'Application.ScreenUpdating = False
'
'On Error Resume Next

    Dim MyWkb As String
    
    Dim MyWkSh As String

    Dim csPath As String
    
    Dim lRw As Long
    
    Dim ws As Worksheet
    
    
    With ActiveWorkbook

        'duplicate the Price-Desc-Cat-Prop65 sheet and eliminate columns B - G,I - T

        .Worksheets("Price-Desc-Cat-Prop65").Copy Before:=.Worksheets("Price-Desc-Cat-Prop65")

        ActiveSheet.Name = "ParentSKUMapping"
        
        Range("B2").Select
        
        ActiveCell.ListObject.Name = "Table1_1"

        'remove unwanted columns

        Columns("I:T").Delete
        Columns("B:G").Delete

        Range("C1").EntireColumn.Insert

        ActiveSheet.Range("C1").Value = "LHAParentSku"
        
        'loop through all cells in column F (ProdName) on the ParentSKUs sheet, match parent SKUs
        
        Set ws = .Worksheets("ParentSKUMapping")
        
        lRw = ActiveSheet.Range("A:A").SpecialCells(xlCellTypeLastCell).Row

        Dim rng As Range
                
        Dim cl As Range
        
        'Dim strFound As String
        
        Dim i As Long
        
        Set rng = ws.Range("B2:B" & lRw)
        
        i = 2
        
        lRw = Worksheets("ParentSKUs").Range("A:A").SpecialCells(xlCellTypeLastCell).Row
        
'        For Each cl In rng
            
        Range("C" & i).Value = "=INDEX(ParentSKUs[[SKU]:[ProdName]],MATCH([@ProdName],ParentSKUs[ProdName],0),1)"
            
'            i = i + 1
        
'        Next cl
    
    End With
    
     'create the export file for the parent SKU mapping information

    MyWkSh = Worksheets("Vendor Info").Range("B2").Value & " AX Parent SKU Mapping Import"

    MyWkb = Format(Now, "YYYY-MM-DD-HHMMSS") & " " & MyWkSh & ".xlsx"

    csPath = Application.ActiveWorkbook.Path & "/"

    Worksheets("ParentSKUMapping").Copy

    With ActiveWorkbook

        Application.DisplayAlerts = False

        'save the workbook

        .SaveAs Filename:=csPath & MyWkb

        'rename the worksheet to Sheet1

        Worksheets("ParentSKUMapping").Name = "Sheet1"

        'remove all formatting after converting table to a range

        Dim rList As Range

        With Worksheets("Sheet1").ListObjects("Table1_1")
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
    
    'create the export file for the actual information

    MyWkSh = Worksheets("Vendor Info").Range("B2").Value & " AX Parent SKU Web Import"

    MyWkb = Format(Now, "YYYY-MM-DD-HHMMSS") & " " & MyWkSh & ".xlsx"

    csPath = Application.ActiveWorkbook.Path & "/"

    Worksheets("ParentSKUs").Copy

    With ActiveWorkbook

        Application.DisplayAlerts = False

        'save the workbook

        .SaveAs Filename:=csPath & MyWkb

        'rename the worksheet to Sheet1

        Worksheets("ParentSKUs").Name = "Sheet1"

        'remove all formatting after converting table to a range

        With Worksheets("Sheet1").ListObjects("ParentSKUs")
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
'
'    Application.ScreenUpdating = True
'
''    MsgBox "The import file was generated for batch upload."
'
'    Worksheets("CommandCentral").Cells(13, 23).Value = Format(Now, "mm/dd/yyyy")
'    Worksheets("CommandCentral").Cells(14, 23).Value = Format(Now, "hh:mm ampm")

End Sub
