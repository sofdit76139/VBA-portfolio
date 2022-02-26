Attribute VB_Name = "mod_CreateAXImportFile2"
Option Explicit

Sub CreateMyAXImport2()

'this sub will add to an existing import file or create it first

'2022-02-08

Dim cl As Range
Dim rng As Range
Dim rngEval As Range
Dim lRw As Long
Dim lRw2 As Long
Dim i As Long
Dim ws As Worksheet
Dim ws2 As Worksheet
Dim curWkb As String
Dim sourceCol As Integer
Dim rowCount As Long
Dim currentRow As Integer
Dim currentRowValue As String
Dim curCalWeek As Integer
Dim MyWkb As String
Dim MyWkSh As String
Dim MyFilePath As String
Dim csPath As String
Dim curWeekUploadFile As Workbook
Dim strUserName As String
Dim fso As New Scripting.FileSystemObject

Application.ScreenUpdating = False

On Error Resume Next

curWkb = "Price Change Template.xlsb"

With Workbooks(curWkb)

    Set ws = .Worksheets("MAPChanges")

    Set ws2 = .Worksheets("AXBatchImport2")

    lRw = ws.Range("A:A").SpecialCells(xlCellTypeLastCell).Row

    lRw2 = ws2.Range("A:A").SpecialCells(xlCellTypeLastCell).Row

    Set rngEval = ws.Range("L2:L" & lRw)

    sourceCol = 1   'column A has a value of 1

    rowCount = Application.WorksheetFunction.CountIf(rngEval, "yes")
    
    'find the first cell in the column that is not empty and copy the value over

    For Each cl In rngEval

        If cl.Value = "yes" Then

            'find the first blank cell and select it
            For currentRow = 2 To rowCount + 1
            
                currentRowValue = ws2.Cells(currentRow, sourceCol).Value
                
                If IsEmpty(currentRowValue) Or currentRowValue = "" Then
                
                    ws2.Cells(currentRow, sourceCol).Offset(0, 1).Value = cl.Offset(0, -1).Value ' copy over the price
                    ws2.Cells(currentRow, sourceCol).NumberFormat = "General"
                    ws2.Cells(currentRow, sourceCol).Offset(0, 2).NumberFormat = "@" ' make sure the item ID is copied over as text
                    ws2.Cells(currentRow, sourceCol).Value = cl.Offset(0, -11).Value ' copy over the item id
                    
                    Exit For
                    
                End If
            Next

        End If

    Next cl
    
End With
    
'add all lines to the current week's batch import file

'check if the workbook already exists; if not, create it

    curCalWeek = WorksheetFunction.WeekNum(Now, vbMonday) 'added to the file name for the week


'see if the workbook exists; if not, create it and add headers

    strUserName = VBA.Interaction.Environ$("UserName")

    MyFilePath = ""

    MyWkb = Format(Now, "YYYY") & " Week " & curCalWeek & " MAP Changes.xlsx"

    csPath = "C:\Users\" & strUserName & "\OneDrive - COMPANY\Merchandising Documents\AX Imports\PricingUpdates\"

    MyFilePath = csPath & MyWkb
    
   If Dir(csPath, vbDirectory) <> "" And Dir(MyFilePath) <> "" Then ' check if one of them is the upload file

        Workbooks.Open (csPath & MyWkb)
    
    Else
    
        Set curWeekUploadFile = Workbooks.Add
        
        With ActiveSheet
            Range("A1").Cells.Value = "ItemId"
            Range("B1").Cells.Value = "LHAMAPPrice"
        End With
        
        curWeekUploadFile.SaveAs Filename:=csPath & MyWkb
    
    End If
    
    'copy all lines from the price list to the batch upload workbook
    
    Set ws = Workbooks(curWkb).Worksheets("AXBatchImport2")

    Set ws2 = Workbooks(MyWkb).Worksheets("Sheet1")

    lRw = ws.Range("A:A").SpecialCells(xlCellTypeLastCell).Row

    lRw2 = ws2.Range("A:A").SpecialCells(xlCellTypeLastCell).Row
    
    sourceCol = 1
    
    'find the first blank cell and select it
    For currentRow = 2 To lRw2 + 1
    
        currentRowValue = Cells(currentRow, sourceCol).Value
        
        If IsEmpty(currentRowValue) Or currentRowValue = "" Then
        
            ws2.Cells(currentRow, sourceCol).Offset(0, 3).NumberFormat = "@"
            ws.Range("A2:B" & lRw).Copy
            ws2.Cells(currentRow, sourceCol).Select
            ws2.Cells(currentRow, sourceCol).PasteSpecial Paste:=xlValues
            
            Exit For
            
        End If
    Next
    
    lRw2 = ws2.Range("A:A").SpecialCells(xlCellTypeLastCell).Row
    
    'remove any duplicates from column ItemId
    With ws2
    
        For Each cl In ws2.Range("A2:A" & lRw2)
                
            cl.Value = WorksheetFunction.Trim(cl.Value)
                
        Next
    
        Application.ErrorCheckingOptions.NumberAsText = False
    
        .Range("A1:B" & lRw2).CurrentRegion.RemoveDuplicates Columns:=1, Header:=xlYes
        
    End With
    
    'remove any blank rows
    Columns("A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    
    
    Workbooks(MyWkb).Save
    Workbooks(MyWkb).Close

    Set curWeekUploadFile = Nothing 'release the object

    'put a timestamp on the command central worksheet

    Worksheets("CommandCentral").Cells(6, 9).Value = Format(Now, "mm/dd/yyyy")
    Worksheets("CommandCentral").Cells(6, 10).Value = Format(Now, "hh:mm ampm")

    Set fso = Nothing
    
    Sheets("CommandCentral").Select
    
    Application.ScreenUpdating = True
    
    MsgBox "The export is now complete.", vbInformation

End Sub

Function DeleteDataConnections(WkbtoSever As String)
'delete all queries and connections in the workbook

    If Workbooks(WkbtoSever).Queries.Count = 0 Then
    
        Exit Function
        
    Else

        For i = 1 To Workbooks(WkbtoSever).Queries.Count
        
            If Workbooks(WkbtoSever).Queries.Count = 0 Then
            
                Exit Function
                
            Else
            
                Workbooks(WkbtoSever).Connections.Item(i).Delete
                
                Workbooks(WkbtoSever).Queries.Item(i).Delete
                
            End If
            
            i = i - 1
        Next i
        
    End If

End Function
