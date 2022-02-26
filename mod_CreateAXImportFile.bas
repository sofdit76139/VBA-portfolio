Attribute VB_Name = "mod_CreateAXImportFile"
Option Explicit

Sub CreateMyAXImport()

'this sub will add to an existing import file for purchase price changes or create it first

'2022-02-09

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
Dim effectiveDate As String
Dim curCalWeek As Integer
Dim MyWkb As String
Dim MyWkSh As String
Dim MyFilePath As String
Dim csPath As String
Dim curWeekUploadFile As Workbook
Dim strUserName As String
Dim myWsName As String
Dim fso As New Scripting.FileSystemObject
Dim myFileRepositoryName() As String
Dim currentBuyer As String
Dim currentVendor As String
Dim myBuyerFile As Workbook
Dim myWsNameNew As String

Application.ScreenUpdating = False

On Error Resume Next

curWkb = "Price Change Template.xlsb"

With Workbooks(curWkb)

    Set ws = .Worksheets("PricingChanges")

    Set ws2 = .Worksheets("AXBatchImport")

    lRw = ws.Range("A:A").SpecialCells(xlCellTypeLastCell).Row

    lRw2 = ws2.Range("A:A").SpecialCells(xlCellTypeLastCell).Row

    Set rngEval = ws.Range("R2:R" & lRw)

    sourceCol = 1   'column A has a value of 1

    rowCount = Application.WorksheetFunction.CountIf(rngEval, "yes")
    
    effectiveDate = .Worksheets("VendorInfo").Range("A5").Value 'set the effective date

    'find the first cell in the column that is not empty and copy the value over

    For Each cl In rngEval

        If cl.Value = "yes" Then

            'find the first blank cell and select it
            For currentRow = 2 To rowCount + 1
            
                currentRowValue = ws2.Cells(currentRow, sourceCol).Value
                
                If IsEmpty(currentRowValue) Or currentRowValue = "" Then
                
                    ws2.Cells(currentRow, sourceCol).Offset(0, 1).Value = cl.Offset(0, -3).Value ' copy over the price
                    ws2.Cells(currentRow, sourceCol).NumberFormat = "General"
                    ws2.Cells(currentRow, sourceCol).Offset(0, 3).NumberFormat = "@" ' make sure the item ID is copied over as text
                    ws2.Cells(currentRow, sourceCol).Offset(0, 3).Value = cl.Offset(0, -17).Value ' copy over the item id
                    ws2.Cells(currentRow, sourceCol).Offset(0, 4).Value = cl.Offset(0, -4).Value ' copy over the unit id
                    ws2.Cells(currentRow, sourceCol).Value = Worksheets("VendorInfo").Range("A2").Value ' copy over the vendor id
                    
                    If IsEmpty(effectiveDate) Then
                        
                        ws2.Cells(currentRow, sourceCol).Offset(0, 2).Value = Format(Now, "mm/dd/yyyy") 'use today's date as effective date
                    
                    Else
                        
                        ws2.Cells(currentRow, sourceCol).Offset(0, 2).Value = effectiveDate
                    
                    End If
                    
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

    MyWkb = Format(Now, "YYYY") & " Week " & curCalWeek & " Price Changes.csv"

    csPath = "C:\Users\" & strUserName & "\OneDrive - COMPANY\Merchandising Documents\AX Imports\PricingUpdates\" 'depends on typical path to OneDrive

    MyFilePath = csPath & MyWkb
    
    MyFilePathAlternate = csPathAlternate & MyWkb
    
    If Dir(csPath, vbDirectory) <> "" And Dir(MyFilePath) <> "" Then ' check if one of them is the upload file

        Workbooks.Open (csPath & MyWkb)
    
    Else
    
        Set curWeekUploadFile = Workbooks.Add
        
        With ActiveSheet
        
            Range("A1").Cells.Value = "VendorId"
            Range("B1").Cells.Value = "Amount"
            Range("C1").Cells.Value = "FromDate"
            Range("D1").Cells.Value = "ItemId"
            Range("E1").Cells.Value = "UnitId"
            
        End With
        
        curWeekUploadFile.SaveAs Filename:=csPath & MyWkb, FileFormat:=xlCSV
    
    End If
    
    'copy all lines from the price list to the batch upload workbook
    
    Set ws = Workbooks(curWkb).Worksheets("AXBatchImport")

    myWsName = fso.GetBaseName(Workbooks(MyWkb).Name)

    Set ws2 = Workbooks(MyWkb).Worksheets(myWsName)

    lRw = ws.Range("A:A").SpecialCells(xlCellTypeLastCell).Row

    lRw2 = ws2.Range("A:A").SpecialCells(xlCellTypeLastCell).Row
    
    sourceCol = 1
    
    'find the first blank cell and select it
    For currentRow = 2 To lRw2 + 1
    
        currentRowValue = Cells(currentRow, sourceCol).Value
        
        If IsEmpty(currentRowValue) Or currentRowValue = "" Then
        
            ws2.Cells(currentRow, sourceCol).Offset(0, 3).NumberFormat = "@"
            ws.Range("A2:E" & lRw).Copy
            ws2.Cells(currentRow, sourceCol).Select
            ws2.Cells(currentRow, sourceCol).PasteSpecial Paste:=xlValues
            
            Exit For
        End If
    Next
    
    lRw2 = ws2.Range("A:A").SpecialCells(xlCellTypeLastCell).Row
 
    ws2.Range("C2:C" & lRw2).NumberFormat = "mm/dd/yyyy"
    ws2.Range("D2:D" & lRw2).NumberFormat = "@"
    
    'remove any duplicates from column ItemId
    With Worksheets(myWsName)
    
        For Each cl In ws2.Range("D2:D" & lRw2)
                
            cl.Value = WorksheetFunction.Trim(cl.Value)
                
        Next
    
        Application.ErrorCheckingOptions.NumberAsText = False
        
        .Range("A1:E" & lRw2).CurrentRegion.Sort Key1:=Range("C1"), Order1:=xlDescending 'sort by date first
    
        .Range("A1:E" & lRw2).CurrentRegion.RemoveDuplicates Columns:=4, Header:=xlYes 'eliminate duplicates
        
    End With
    
    'remove any blank rows
    Columns("A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    
    Workbooks(MyWkb).Save
    Workbooks(MyWkb).Close

    Set curWeekUploadFile = Nothing 'release the object

    'copy all items you need to contact the buyer about into the changelog workbook, along with a list of vendors you have processed
    'check if the file exists first; if it does, open it, if not, create it

    strUserName = VBA.Interaction.Environ$("UserName")

    MyFilePath = ""

'    MyWkb = Format(Now, "YYYY") & " Week " & curCalWeek & " Change Log.xlsx" 'for the year and the week; changed on 1/28/2022 to just the year
    
    MyWkb = Format(Now, "YYYY") & " Purchase Price Updates Change Log.xlsx"

    csPath = "C:\Users\" & strUserName & "\OneDrive - COMPANY\Merchandising Documents\AX Imports\PricingUpdates\"

    MyFilePath = csPath & MyWkb

    If Dir(MyFilePath) <> "" And Dir(MyFilePath) = MyWkb Then ' check if one of them is the change log file

        Set curWeekUploadFile = Workbooks.Open(MyFilePath)

        Set curWeekUploadFile = Workbooks(MyWkb)

         'check if the sheet exists, and if it does, delete it to avoid duplicates

        Set ws = Nothing ' release the object

        Set ws = Workbooks(curWkb).Worksheets("PricingChanges")

        myWsName = ws.Cells(2, 4).Value & "-" & ws.Cells(2, 5).Value 'the worksheet is named after the vendor and the buyer

        For Each ws In curWeekUploadFile.Worksheets
        
            If ws.Name = myWsName Then

                MyWkSh = ws.Name

                Application.DisplayAlerts = False

                curWeekUploadFile.Worksheets(MyWkSh).Delete

                Application.DisplayAlerts = True

            End If

        Next ws

    Else ' if not, create the file

        Set curWeekUploadFile = Workbooks.Add

        curWeekUploadFile.SaveAs Filename:=MyFilePath

        Set curWeekUploadFile = Workbooks(MyWkb)

        Set ws2 = curWeekUploadFile.Worksheets("Sheet1")

'        ws2.Name = "Week " & curCalWeek & " Change Log"

        ws2.Name = "Change Log"

        Set ws2 = curWeekUploadFile.Worksheets("Change Log")

        With ActiveSheet
        
            Range("A1").Cells.Value = "VendorID"
            Range("B1").Cells.Value = "VendorName"
            Range("C1").Cells.Value = "ProcessedDate"
            Range("D1").Cells.Value = "EffectiveDate"
            
        End With

        'format the range as a table
        ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A1:$D2"), , xlYes).Name = _
        "VendorLog"
        ActiveSheet.ListObjects("VendorLog").TableStyle = "TableStyleMedium9"

    End If

    'write the vendor information to the change log tab

    Set ws = Workbooks(curWkb).Worksheets("VendorInfo")

'    Set ws2 = curWeekUploadFile.Worksheets("Week " & curCalWeek & " Change Log")

    Set ws2 = curWeekUploadFile.Worksheets("Change Log")

    lRw2 = ws2.Range("A:A").SpecialCells(xlCellTypeLastCell).Row

    sourceCol = 1

    'find the first blank cell and select it
    For currentRow = 2 To lRw2 + 1
    
        currentRowValue = Cells(currentRow, sourceCol).Value
        
        If IsEmpty(currentRowValue) Or currentRowValue = "" Then
        
            ws.Range("A2:B2").Copy
            ws2.Activate
            ws2.Cells(currentRow, sourceCol).Select
            ws2.Cells(currentRow, sourceCol).PasteSpecial Paste:=xlValues
            ws2.Cells(currentRow, sourceCol).Offset(0, 2).Value = Format(Now, "mm/dd/yyyy")
            ws2.Activate
            ws2.Cells(currentRow, sourceCol).Offset(0, 3).Select
            ws2.Cells(currentRow, sourceCol).Offset(0, 3).NumberFormat = "mm/dd/yyyy"
            ws.Range("A5").Copy
            ws2.Cells(currentRow, sourceCol).Offset(0, 3).PasteSpecial Paste:=xlValues 'paste the effective date of the purchase price list
            MsgBox ("Vendor was added to change log."), vbInformation
            
            Exit For
            
        End If
        
    Next currentRow

    lRw2 = ws2.Range("A:A").SpecialCells(xlCellTypeLastCell).Row

    'remove any duplicates from column VendorID
    With ws2

        For Each cl In ws2.Range("A2:A" & lRw2)

            cl.Value = WorksheetFunction.Trim(cl.Value)

        Next

        With ws2.Sort
        
            .SortFields.Add Key:=Range("A1"), Order:=xlAscending
            .SortFields.Add Key:=Range("D1"), Order:=xlDescending
            .SetRange Range("A1:D" & lRw2)
            .Header = xlYes
            .Apply
            
        End With
        
        .Range("A1:D" & lRw2).CurrentRegion.RemoveDuplicates Columns:=Array(1, 4), Header:=xlYes ' kick out duplicates based on effective date

    End With
    
    Set ws = Nothing ' release the object

    Set ws = Workbooks(curWkb).Worksheets("PricingChanges")

    myWsName = ws.Cells(2, 4).Value & "-" & ws.Cells(2, 5).Value 'the worksheet is named after the vendor and the buyer

    'copy over the pricing changes sheet and rename it

    ws.Copy After:=curWeekUploadFile.Sheets(Sheets.Count)

    ActiveSheet.Name = myWsName

    'delete all data connections in the change log workbook
    DeleteDataConnections (curWeekUploadFile.Name)

    'last step: bundle all of the worksheets for each buyer into their separate files
    'check if the file exists first; if it does, open it, if not, create it

    myFileRepositoryName = Split(myWsName, "-")

    currentVendor = myFileRepositoryName(0) 'pass the current vendor to the variable
    
    currentBuyer = myFileRepositoryName(1)  'pass the current buyer to the variable

    strUserName = VBA.Interaction.Environ$("UserName")

    MyFilePath = ""

    MyWkb = Format(Now, "YYYY") & " Week " & curCalWeek & " " & currentBuyer & ".xlsx"

    csPath = "C:\Users\" & strUserName & "\OneDrive - COMPANY\Merchandising Documents\AX Imports\PricingUpdates\"

    MyFilePath = csPath & MyWkb

    If Dir(MyFilePath) <> "" Then ' check if one of them is the buyer's file

        Set myBuyerFile = Workbooks.Open(MyFilePath)

        Set myBuyerFile = Workbooks(MyWkb)

        'check if the file has a sheet for the vendor already; if yes, delete it to avoid duplicates

        myWsNameNew = currentVendor  'the worksheet is named after the vendor and the buyer

            myBuyerFile.Activate

            Set ws2 = Nothing ' release the object

            For Each ws2 In myBuyerFile.Worksheets
            
                If ws2.Name = myWsNameNew And myBuyerFile.Worksheets.Count > 1 Then

                    MyWkSh = ws2.Name

                    Application.DisplayAlerts = False

                    myBuyerFile.Worksheets(MyWkSh).Delete

                    Application.DisplayAlerts = True
                    
                Else
                
                    Worksheets.Add.Name = "Sheet1"
                    
                    MyWkSh = ws2.Name

                    Application.DisplayAlerts = False

                    myBuyerFile.Worksheets(MyWkSh).Delete

                    Application.DisplayAlerts = True

                End If

            Next ws2

        Set ws = Nothing ' release the object

    Else ' if not, create the file

        Set myBuyerFile = Workbooks.Add

        myBuyerFile.SaveAs Filename:=MyFilePath

        Set myBuyerFile = Workbooks(MyWkb)

    End If

    Set ws = curWeekUploadFile.Worksheets(myWsName)

    'copy over the pricing changes sheet and rename it

    myBuyerFile.Activate

    ws.Copy After:=myBuyerFile.Sheets(Sheets.Count)

    Set ws2 = Nothing ' release the object

    Set ws2 = myBuyerFile.Worksheets(myWsName)

    myWsNameNew = currentVendor

    ws2.Name = myWsNameNew

    'if there is still a Sheet1 in the workbook, delete it

    For Each ws2 In myBuyerFile.Worksheets
    
        If ws2.Name = "Sheet1" Then

            MyWkSh = ws2.Name

            Application.DisplayAlerts = False

            myBuyerFile.Worksheets(MyWkSh).Delete

            Application.DisplayAlerts = True

        End If

    Next ws2

    'save the buyer file
    myBuyerFile.Save
    myBuyerFile.Close

    'delete all buyer sheets in the changelog file for this buyer

    Set ws = Nothing ' release the object

    For Each ws In curWeekUploadFile.Worksheets
    
        If ws.Name = myWsName Then

            MyWkSh = ws.Name

            Application.DisplayAlerts = False

            'On Error Resume Next

            curWeekUploadFile.Worksheets(MyWkSh).Delete

            Application.DisplayAlerts = True

        End If

    Next ws

    curWeekUploadFile.Worksheets("Change Log").Sort.SortFields.Clear

    'save and close the changelog file
    curWeekUploadFile.Save
    curWeekUploadFile.Close

    'put a timestamp on the command central worksheet

    Worksheets("CommandCentral").Cells(2, 9).Value = Format(Now, "mm/dd/yyyy")
    Worksheets("CommandCentral").Cells(2, 10).Value = Format(Now, "hh:mm ampm")

    Set fso = Nothing
    
    Sheets("CommandCentral").Select
    
    Application.ScreenUpdating = True
    
    MsgBox "The export is now complete.", vbInformation

End Sub

Function DeleteDataConnections(WkbtoSever As String)
'delete all queries and connections in the workbook

Dim i As Long

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
