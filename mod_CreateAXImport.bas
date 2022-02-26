Attribute VB_Name = "mod_CreateAXImport"
Option Explicit

Sub CreateAXImportCurrentYear()
'creates a SKU import file for AX by vendor

'2022-01-27

Dim MyWkbSrc As String
Dim MyWkbTrg As String
Dim MyWkb As String
Dim MyWkSh As String
Dim csPath As String
Dim csPath2 As String
Dim lRw As Long
Dim curWeekUploadFile As Workbook
Dim ws As Worksheet
Dim ws2 As Worksheet
Dim wkB As Workbook
Dim cl As Range
Dim rng As Range
Dim rngEval As Range
Dim i As Long
Dim j As Long
Dim myVendorIDList() As Variant
Dim outputArrayIDs() As String
Dim outputArrayNames() As String
Dim myOptionButton As OLEObject
Dim myOptionButton2 As OLEObject
Dim myOptName As String
Dim strUserName As String
Dim fso As Object
Dim xRet As Boolean
Dim myItem As Variant
Dim vendorID As String
Dim vendorName As String
Dim MyFileName As String
Dim MyRangeName As String
Dim curCalWeek As Integer
Dim MyFilePath As String
Dim sourceCol As Long
Dim rowCount As Long
Dim currentRow As Long
Dim currentRowValue As String
Dim lRw3 As Long
Dim lCol As Long
Dim lColLetter As String
Dim myFileName2 As String
Dim myMessageBoxResult As String
Dim myBuyerGroupID As String

    Application.ScreenUpdating = False

    On Error Resume Next

'**************************************************************************************************************
'determine user the script should be run for
'**************************************************************************************************************

    Set ws = Workbooks("create_SKU_import_files.xlsb").Worksheets("CommandCentral")
    
    Workbooks("create_SKU_import_files.xlsb").Activate
    
    ws.Select
    
    Set myOptionButton = ws.OLEObjects("optPERSONA")
    Set myOptionButton2 = ws.OLEObjects("optPERSONB")
    
    If myOptionButton.Object.Value = True Then
    
        myOptName = "Cameron"
        
    ElseIf myOptionButton2.Object.Value = True Then
    
        myOptName = "Kimberly"
        
    End If

'**************************************************************************************************************
'delete any accidental/unwanted entries in the template workbooks
'**************************************************************************************************************
                
    strUserName = VBA.Interaction.Environ$("UserName")
    
    csPath = "C:\Users\" & strUserName & "\OneDrive - COMPANY\Merchandising Documents\AX Imports\"
    
    MyFileName = csPath & "CreatedSKUs2PA.xlsx"
    
    MyRangeName = "A:Q"
    
    DeleteUnwantedEntries MyFileName, MyRangeName
    
    MyFileName = csPath & "CreatedSKUsUploadPricing.xlsx"
    
    MyRangeName = "A:Q"
    
    DeleteUnwantedEntries MyFileName, MyRangeName
    
    MyFileName = ""

    csPath = ""
    
    strUserName = ""

'**************************************************************************************************************
'copy over the latest version of the original file, current year
'**************************************************************************************************************
    
    strUserName = VBA.Interaction.Environ$("UserName")

    csPath = "C:\Users\" & strUserName & "\OneDrive - COMPANY\Merchandising Documents\AX Imports\New SKUs\"
    
    csPath2 = "C:\Users\" & strUserName & "\OneDrive - COMPANY\Merchandising Documents\Kidron Merchandising\"

    MyWkb = Year(Now) & " " & myOptName & "'s New Item List.xlsx"

    MyWkbSrc = csPath2 & MyWkb
    
    'if the source workbook does not exist, abort
    
    MyFilePath = ""
    
    MyFilePath = Dir(csPath2 & MyWkb)

    If MyFilePath = "" Then
        
        MsgBox "The source file " & MyWkb & " is missing. This script will now exit.", vbCritical, "File Missing!"
        
        Exit Sub
        
    End If
    
    ' have the user check first if all information needed to execute the code is there

    myMessageBoxResult = MsgBox("Have you double-checked the information in the New Item List?" & vbCrLf & vbCrLf & "The script will abort if information in the following fields is missing:" _
        & vbCrLf _
        & vbCrLf & "- Product Name" _
        & vbCrLf & "- Lowest Category" _
        & vbCrLf & "- Purchase Unit" _
        & vbCrLf & "- Selling Unit" _
        & vbCrLf & "- Buyer #" _
        & vbCrLf & "- Vendor ID" _
        & vbCrLf & "- Vendor Name" _
        & vbCrLf & "- Cost" _
        & vbCrLf & "- Retail Price" _
        & vbCrLf & "- External Item #" _
        , vbOKCancel + vbQuestion, "Mandatory Fields:")
    
    If myMessageBoxResult = vbCancel Then
    
        Exit Sub
        
    End If

    MyWkbTrg = csPath & "SKU_working_file.xlsx"
    
    'check if the working file is still open; if yes, close it
    
    xRet = IsWorkBookOpen("SKU_working_file.xlsx") 'INCLUDE FULL NAME OF WORKBOOK HERE

    If xRet Then

        Application.DisplayAlerts = False

        Workbooks("SKU_working_file.xlsx").Close

        Application.DisplayAlerts = True

    End If

    Set fso = CreateObject("scripting.filesystemobject")
    
    fso.DeleteFile MyWkbTrg
    
    fso.CopyFile Source:=MyWkbSrc, Destination:=MyWkbTrg
    
    Set fso = Nothing

'check if the workbook is already open; if not, open it
    
    Workbooks.Open Filename:=MyWkbTrg
    
    Set wkB = Workbooks("SKU_working_file.xlsx")
    
    Set ws = wkB.Worksheets("Shelley")
    
    ws.Columns.EntireColumn.Hidden = False 'if somebody has hidden any columns, unhide all

    lRw = ws.Range("A:A").Find(what:="X", after:=ws.Range("A3"), LookIn:=xlValues, LookAt:=xlWhole).Row - 1 'find the last row the delimiter X was entered into, go up one row, use that value as last row
    
    If lRw = 0 Then
    
        Application.DisplayAlerts = False
        Workbooks("SKU_working_file.xlsx").Close
        Application.DisplayAlerts = True
    
        MsgBox "No delimiter was found. Please check column A. This script will now exit.", vbCritical, "Information Missing!"
    
    End If
    
With wkB

'    ' check if all important fields have entries; if no, abort with message

    For Each cl In ws.Range("A1:A" & lRw)

            'loop only through the empty cells until you find one that fits the criteria

                If IsEmpty(cl) And cl.Interior.Color = RGB(248, 203, 173) Then

                        If IsEmpty(cl.Offset(0, 1)) Then ' product name
                            MsgBox "Please check the information in cell " & cl.Offset(0, 1).Address & "! This script will now exit.", vbCritical, "Information Missing!"
                            Application.DisplayAlerts = False
                            Workbooks("SKU_working_file.xlsx").Close
                            Application.DisplayAlerts = True
                            Exit Sub

                        ElseIf IsEmpty(cl.Offset(0, 6)) Then ' lowest category
                            MsgBox "Please check the information in cell " & cl.Offset(0, 6).Address & "! This script will now exit.", vbCritical, "Information Missing!"
                            Application.DisplayAlerts = False
                            Workbooks("SKU_working_file.xlsx").Close
                            Application.DisplayAlerts = True
                            Exit Sub

                        ElseIf IsEmpty(cl.Offset(0, 14)) Then ' purchase unit
                            MsgBox "Please check the information in cell " & cl.Offset(0, 14).Address & "! This script will now exit.", vbCritical, "Information Missing!"
                            Application.DisplayAlerts = False
                            Workbooks("SKU_working_file.xlsx").Close
                            Application.DisplayAlerts = True
                            Exit Sub

                        ElseIf IsEmpty(cl.Offset(0, 15)) Then ' selling unit
                            MsgBox "Please check the information in cell " & cl.Offset(0, 15).Address & "! This script will now exit.", vbCritical, "Information Missing!"
                            Application.DisplayAlerts = False
                            Workbooks("SKU_working_file.xlsx").Close
                            Application.DisplayAlerts = True
                            Exit Sub

                        ElseIf IsEmpty(cl.Offset(0, 16)) Then ' buyer ID
                            MsgBox "Please check the information in cell " & cl.Offset(0, 16).Address & "! This script will now exit.", vbCritical, "Information Missing!"
                            Application.DisplayAlerts = False
                            Workbooks("SKU_working_file.xlsx").Close
                            Application.DisplayAlerts = True
                            Exit Sub
                            
                        ElseIf Not (cl.Offset(0, 16) Like "Buyer ?") Then ' buyer ID wrong
                            MsgBox "Please check the information in cell " & cl.Offset(0, 16).Address & "! You need to enter the buyer ID. This script will now exit.", vbCritical, "Information Wrong!"
                            Application.DisplayAlerts = False
                            Workbooks("SKU_working_file.xlsx").Close
                            Application.DisplayAlerts = True
                            Exit Sub

                        ElseIf IsEmpty(cl.Offset(0, 17)) Then ' vendor ID
                            MsgBox "Please check the information in cell " & cl.Offset(0, 17).Address & "! This script will now exit.", vbCritical, "Information Missing!"
                            Application.DisplayAlerts = False
                            Workbooks("SKU_working_file.xlsx").Close
                            Application.DisplayAlerts = True
                            Exit Sub
                            
                        ElseIf Not (cl.Offset(0, 17) Like "V?????") Then ' vendor ID wrong
                            MsgBox "Please check the information in cell " & cl.Offset(0, 17).Address & "! The vendor ID may be wrong. This script will now exit.", vbCritical, "Information Wrong!"
                            Application.DisplayAlerts = False
                            Workbooks("SKU_working_file.xlsx").Close
                            Application.DisplayAlerts = True
                            Exit Sub

                        ElseIf IsEmpty(cl.Offset(0, 18)) Then ' vendor name
                            MsgBox "Please check the information in cell " & cl.Offset(0, 18).Address & "! This script will now exit.", vbCritical, "Information Missing!"
                            Application.DisplayAlerts = False
                            Workbooks("SKU_working_file.xlsx").Close
                            Application.DisplayAlerts = True
                            Exit Sub

                        ElseIf IsEmpty(cl.Offset(0, 19)) Then ' cost/purchase price
                            MsgBox "Please check the information in cell " & cl.Offset(0, 19).Address & "! This script will now exit.", vbCritical, "Information Missing!"
                            Application.DisplayAlerts = False
                            Workbooks("SKU_working_file.xlsx").Close
                            Application.DisplayAlerts = True
                            Exit Sub

                        ElseIf IsEmpty(cl.Offset(0, 20)) Then ' standard cost
                            MsgBox "Please check the information in cell " & cl.Offset(0, 20).Address & "! This script will now exit.", vbCritical, "Information Missing!"
                            Application.DisplayAlerts = False
                            Workbooks("SKU_working_file.xlsx").Close
                            Application.DisplayAlerts = True
                            Exit Sub

                        ElseIf IsEmpty(cl.Offset(0, 24)) Then ' retail price
                            MsgBox "Please check the information in cell " & cl.Offset(0, 24).Address & "! This script will now exit.", vbCritical, "Information Missing!"
                            Application.DisplayAlerts = False
                            Workbooks("SKU_working_file.xlsx").Close
                            Application.DisplayAlerts = True
                            Exit Sub

                        ElseIf IsEmpty(cl.Offset(0, 31)) Then ' external item ID
                            MsgBox "Please check the information in cell " & cl.Offset(0, 31).Address & "! This script will now exit.", vbCritical, "Information Missing!"
                            Application.DisplayAlerts = False
                            Workbooks("SKU_working_file.xlsx").Close
                            Application.DisplayAlerts = True
                            Exit Sub

                    End If

                End If

            Next
    
'**************************************************************************************************************
'process the changes that are ready to go:
'**************************************************************************************************************
    
'this worked at first, then it started crashing the file; turning off for now
''turn off OneDrive sync
'ManageOnedriveSync (1)
    
'**************************************************************************************************************
'1. loop through the cells in the first column to check which ones need to be processed
'2.a find the vendor ID for those, store it
'**************************************************************************************************************
    
'  in the SKU working file, check if the cell has the peach color and is otherwise empty
    
    i = 1
    
    With ws
        
        'find a vendor we don't have SKUs created for, yet; store the ID and name in an array, then loop through the array
        
        For Each cl In .Range("A3:A" & lRw)

            If IsEmpty(cl) And cl.Interior.Color = RGB(248, 203, 173) Then   'loop only through the empty cells until you find one that fits the criteria

                'read the vendor ID in column R, store in variable

                vendorID = cl.Offset(0, 17).Value

                ReDim Preserve myVendorIDList(i)
                
                myVendorIDList(i) = vendorID
                
                i = i + 1
                
            End If

        Next
            
    'if the array is empty, abort and alert the user
            
        If IsArrayAllocated(myVendorIDList) = False Then
        
            Application.DisplayAlerts = False
        
            Workbooks("SKU_working_file.xlsx").Close SaveChanges:=False
            
            Application.DisplayAlerts = True
            
            Windows("create_SKU_import_files.xlsb").Activate
            
            MsgBox "There were no files to be created.", vbExclamation
            
'            'turn on OneDrive sync
'            ManageOnedriveSync (0)
            
            Exit Sub
            
        Else
        
          outputArrayIDs = ArrayRemoveDups(myVendorIDList)
          
          'remove the first empty element from the array
          
          For i = 1 To UBound(outputArrayIDs)
            outputArrayIDs(i - 1) = outputArrayIDs(i)
          Next i
          ReDim Preserve outputArrayIDs(UBound(outputArrayIDs) - 1)
          
'**************************************************************************************************************
'OVERALL LOOP PER VENDOR START
'**************************************************************************************************************
    
          For Each myItem In outputArrayIDs 'loop through all of the vendor IDs in the array
        
            vendorID = myItem
                
            For Each cl In .Range("A1:A" & lRw)
            
            'loop only through the empty cells until you find one that fits the criteria
    
                If IsEmpty(cl) And cl.Interior.Color = RGB(248, 203, 173) And cl.Offset(0, 17).Value = vendorID Then
    
                    'read the vendor name
    
                    vendorName = cl.Offset(0, 18).Value ' store the vendor name in a variable
                    
                    myBuyerGroupID = cl.Offset(0, 16).Value ' store the buyer ID in a variable
    
                    Exit For
    
                End If
    
            Next
            
'**************************************************************************************************************
'2.b use those vendor IDs to create the file/file name for the respective import file
'**************************************************************************************************************
            
        'create a workbook with the week number, vendor ID, vendor name in the file name

            csPath = "C:\Users\" & strUserName & "\OneDrive - COMPANY\Merchandising Documents\AX Imports\New SKUs\"
            
            curCalWeek = WorksheetFunction.WeekNum(Now, vbMonday) 'add week number to the file name for the week
            
            MyFileName = Year(Now) & " Week " & curCalWeek & " Upload" & " - " & vendorID & ".xlsx"
        
        
            'see if the workbook exists; if not, create it and add headers; if it does, create it again and save it over the other one
        
             MyFilePath = ""
        
             MyFilePath = Dir(csPath & MyFileName)
        
             If MyFilePath = "" Then
        
                 Set curWeekUploadFile = Workbooks.Add
        
                 curWeekUploadFile.SaveAs Filename:=csPath & MyFileName
        
                 Set curWeekUploadFile = ActiveWorkbook
        
             Else
                 
                 Set curWeekUploadFile = Workbooks.Add
                 
                 Application.DisplayAlerts = False
        
                 curWeekUploadFile.SaveAs Filename:=csPath & MyFileName
                 
                 Application.DisplayAlerts = True
        
                 Set curWeekUploadFile = ActiveWorkbook
        
             End If
        
            'copy over the header row from the working skus
        
            ws.Range("A2").EntireRow.Copy
        
            curWeekUploadFile.Worksheets("Sheet1").Range("A1").Select
        
            curWeekUploadFile.Worksheets("Sheet1").Range("A1").PasteSpecial Paste:=xlPasteValues
    
'**************************************************************************************************************
'3. loop only through the entries for the vendors previously identified, create the import files by copying over the information
'**************************************************************************************************************

            'loop through the cells for the vendor and copy over the information
        
            Set ws2 = curWeekUploadFile.Worksheets("Sheet1")
        
            sourceCol = 1 'index of column A is 1
        
            Set rngEval = ws.Range("R3:R" & lRw)
        
            rowCount = Application.WorksheetFunction.CountIf(rngEval, vendorID) + 1
        
            'find the first blank cell and select it
        
            For Each cl In rngEval
        
                If cl.Value = vendorID And cl.Offset(0, -17).Value = "" And cl.Offset(0, -17).Interior.Color = RGB(248, 203, 173) Then 'copy only from rows where the first cell is blank and the vendor ID matches what we have previously found
        
                cl.EntireRow.Copy
        
                    'find the first blank cell in the target spreadsheet and copy over the information
                    For currentRow = 2 To rowCount
        
                    currentRowValue = Cells(currentRow, sourceCol).Offset(0, 1).Value
                    
                    If IsEmpty(currentRowValue) Or currentRowValue = "" Then
                    
                        Cells(currentRow, sourceCol).Select
                        
                        Cells(currentRow, sourceCol).PasteSpecial Paste:=xlPasteValues
                        
                        Exit For
                        
                    End If
                    
                    Next currentRow
        
                End If
        
            Next cl

            'make sure that cost/price (standard cost) column is rounded to 2 digits
            
            With curWeekUploadFile
        
                Set ws2 = .Worksheets("Sheet1")
            
                lRw3 = ws2.Range("B:B").SpecialCells(xlCellTypeLastCell).Row
                
                For Each cl In ws2.Range("U2:U" & lRw3)
                
                    cl.Value = WorksheetFunction.Round(cl.Value, 2)
                
                Next cl
            
            End With

            'remove any columns with an "X" or a blank in the header row
        
                With curWeekUploadFile
        
                    lCol = .Worksheets("Sheet1").Range("1:1").SpecialCells(xlCellTypeLastCell).Column 'find the very last column in the spreadsheet
        
                    For i = lCol To 1 Step -1
                    
                        Cells(1, i).Value = WorksheetFunction.Trim(Cells(1, i).Value) 'eliminate unwanted spaces in header
        
                        If Cells(1, i).Value = "X" Or Cells(1, i).Value = "" Then
        
                            Cells(1, i).EntireColumn.Delete
        
                        End If
        
                    Next i
        
                End With
                
            Application.DisplayAlerts = False
        
            curWeekUploadFile.Close SaveChanges:=True
            
            Application.DisplayAlerts = True
            
            Application.Wait (Now + TimeValue("00:00:02"))
    
'**************************************************************************************************************
'4. create the price update workbook
'**************************************************************************************************************

            'check if the template exists in the same directory; if it doesn't, display a warning to the user and abort
            
            csPath = "C:\Users\" & strUserName & "\OneDrive - COMPANY\Merchandising Documents\AX Imports\"
            
            csPath2 = "C:\Users\" & strUserName & "\OneDrive - COMPANY\Merchandising Documents\AX Imports\New SKUs\"
        
            MyFileName = "CreatedSKUsUploadPricing.xlsx"
            
            myFileName2 = Year(Now) & " Week " & curCalWeek & " Pricing Setup" & " - " & vendorID & ".xlsx"
            
            'see if the workbook exists; if not, create it and add headers
    
            MyFilePath = ""
    
            MyFilePath = Dir(csPath & MyFileName)
    
            If MyFilePath = "" Then
                
                GoTo NothingToSeeHere
    
            Else
    
                Workbooks.Open (csPath & MyFileName)
                
                Set curWeekUploadFile = ActiveWorkbook
                
                'save the template as the new name
                
                Application.DisplayAlerts = False
                
                curWeekUploadFile.SaveAs Filename:=csPath2 & myFileName2
                
                Application.DisplayAlerts = True
    
            End If

            'loop through the cells for the vendor and copy over the information
        
                Set ws2 = curWeekUploadFile.Worksheets("ItemInfo")
                
                'then copy over the information
        
                sourceCol = 1 'index of column A is 1
        
                Set rngEval = ws.Range("R3:R" & lRw)
        
                rowCount = Application.WorksheetFunction.CountIf(rngEval, vendorID) + 1
                
'                rowCount = rowCount * 3 'we run this three times, for all three price types
        
                'find the first blank cell and select it
        
                For Each cl In rngEval
        
                    If cl.Value = vendorID And cl.Offset(0, -17).Value = "" And cl.Offset(0, -17).Interior.Color = RGB(248, 203, 173) Then 'copy only from rows where the first cell is blank and the vendor ID matches what we have previously found
                        
                        'find the first blank cell in the target spreadsheet and copy over the information
                        For currentRow = 2 To rowCount
        
                        currentRowValue = ws2.Cells(currentRow, sourceCol).Value
                        
                        If IsEmpty(currentRowValue) Or currentRowValue = "" Then
                            
                            'purchase price
                            ws2.Cells(currentRow, sourceCol).NumberFormat = "@"
                            ws2.Cells(currentRow, sourceCol).Value = cl.Offset(0, 14).Value ' copy over the external item ID
                            ws2.Cells(currentRow, sourceCol).Offset(0, 1).Value = "0"
                            ws2.Cells(currentRow, sourceCol).Offset(0, 3).Value = cl.Offset(0, -3).Value ' copy over the unit ID
                            ws2.Cells(currentRow, sourceCol).Offset(0, 5).Value = "0"
                            ws2.Cells(currentRow, sourceCol).Offset(0, 7).Value = "0"
                            ws2.Cells(currentRow, sourceCol).Offset(0, 9).Value = "AllBlank"
                            ws2.Cells(currentRow, sourceCol).Offset(0, 10).Value = cl.Offset(0, 2).Value ' copy over the purchase price
                            ws2.Cells(currentRow, sourceCol).Offset(0, 11).Value = "USD"
                            ws2.Cells(currentRow, sourceCol).Offset(0, 12).Value = vendorID ' copy over the vendor ID
                            ws2.Cells(currentRow, sourceCol).Offset(0, 14).Value = Format(Now, "MM/DD/YYYY") ' use today's date as "from" date
                            ws2.Cells(currentRow, sourceCol).Offset(0, 16).Value = "2" ' Module 2 is for purchase price
                            
                            Exit For
                            
                        End If
                        
                        Next currentRow
                        
                    End If
        
                Next cl

            'remove any duplicates
            With ws2
            
                For Each cl In ws2.Range("A2:A" & rowCount)
                        
                    cl.Value = WorksheetFunction.Trim(cl.Value)
                        
                Next
            
'                .Range("A1:Q" & rowCount).CurrentRegion.RemoveDuplicates Columns:=Array(1, 17), Header:=xlYes
                .Range("A1:A" & rowCount).CurrentRegion.RemoveDuplicates Columns:=1, Header:=xlYes
                
            End With
        
            'hide the sheet "ItemInfo"
            curWeekUploadFile.Worksheets("ItemInfo").Visible = xlSheetHidden
        
            Application.DisplayAlerts = False
        
            curWeekUploadFile.Close SaveChanges:=True
            
            Application.DisplayAlerts = True
            
            Application.Wait (Now + TimeValue("00:00:02"))

'**************************************************************************************************************
'5. create the buyer update workbook
'**************************************************************************************************************

            'check if the template exists in the same directory; if it doesn't, display a warning to the user and abort
            
            csPath = "C:\Users\" & strUserName & "\OneDrive - COMPANY\Merchandising Documents\AX Imports\"
            
            csPath2 = "C:\Users\" & strUserName & "\OneDrive - COMPANY\Merchandising Documents\AX Imports\New SKUs\"
        
            MyFileName = "CreatedSKUs2PA.xlsx"
            
            myFileName2 = Year(Now) & " Week " & curCalWeek & " - " & myBuyerGroupID & " - " & vendorID & ".xlsx"
            
            'see if the workbook exists; if not, create it and add headers
    
            MyFilePath = ""
    
            MyFilePath = Dir(csPath & MyFileName)
    
            If MyFilePath = "" Then
    
                GoTo NothingToSeeHere2
    
            Else
    
                Workbooks.Open (csPath & MyFileName)
                
                Set curWeekUploadFile = ActiveWorkbook
                
                Set ws2 = curWeekUploadFile.Worksheets("ItemInfo")
                
                'save the template as the new name
                
                Application.DisplayAlerts = False
                
                curWeekUploadFile.SaveAs Filename:=csPath2 & myFileName2
                
                Application.DisplayAlerts = True
    
            End If

            'loop through the cells for the vendor and copy over the information
        
                sourceCol = 1 'index of column A is 1
        
                Set rngEval = ws.Range("R3:R" & lRw)
        
                rowCount = Application.WorksheetFunction.CountIf(rngEval, vendorID) + 1
        
                'find the first blank cell and select it
        
                For Each cl In rngEval
        
                    If cl.Value = vendorID And cl.Offset(0, -17).Value = "" And cl.Offset(0, -17).Interior.Color = RGB(248, 203, 173) Then 'copy only from rows where the first cell is blank and the vendor ID matches what we have previously found
                        
                        'find the first blank cell in the target spreadsheet and copy over the information
                        For currentRow = 2 To rowCount
        
                        currentRowValue = ws2.Cells(currentRow, sourceCol).Value
                        
                        If IsEmpty(currentRowValue) Or currentRowValue = "" Then

                            ws2.Cells(currentRow, sourceCol).Value = cl.Offset(0, -16).Value ' item description
                            ws2.Cells(currentRow, sourceCol).Offset(0, 1).Value = cl.Offset(0, 56).Value ' order quantity for Kidron
                            ws2.Cells(currentRow, sourceCol).Offset(0, 2).Value = cl.Offset(0, 57).Value ' order quantity for Dalton
                            ws2.Cells(currentRow, sourceCol).Offset(0, 3).NumberFormat = "@"
                            ws2.Cells(currentRow, sourceCol).Offset(0, 3).Value = cl.Offset(0, 14).Value ' external item ID
                            ws2.Cells(currentRow, sourceCol).Offset(0, 4).NumberFormat = "@"
                            ws2.Cells(currentRow, sourceCol).Offset(0, 4).Value = cl.Offset(0, 21).Value ' UPC
                            ws2.Cells(currentRow, sourceCol).Offset(0, 5).Value = cl.Offset(0, 7).Value ' retail price
                            ws2.Cells(currentRow, sourceCol).Offset(0, 6).Value = cl.Offset(0, 3).Value ' standard cost
                            ws2.Cells(currentRow, sourceCol).Offset(0, 7).Value = cl.Offset(0, 2).Value ' purcase price
                            ws2.Cells(currentRow, sourceCol).Offset(0, 8).Value = vendorID ' enter the vendor ID
                            ws2.Cells(currentRow, sourceCol).Offset(0, 9).Value = cl.Offset(0, 1).Value ' vendor name
                            ws2.Cells(currentRow, sourceCol).Offset(0, 10).Value = myBuyerGroupID ' enter the buyer ID
                            ws2.Cells(currentRow, sourceCol).Offset(0, 11).Value = cl.Offset(0, 17).Value ' minimum order quantity
                            ws2.Cells(currentRow, sourceCol).Offset(0, 12).Value = cl.Offset(0, 16).Value ' number of items per case
                            ws2.Cells(currentRow, sourceCol).Offset(0, 13).Value = cl.Offset(0, 33).Value ' minimum quantity Kidron
                            ws2.Cells(currentRow, sourceCol).Offset(0, 14).Value = cl.Offset(0, 27).Value ' minimum quantity Dalton
                            ws2.Cells(currentRow, sourceCol).Offset(0, 15).Value = cl.Offset(0, 59).Value ' receipt date
                            ws2.Cells(currentRow, sourceCol).Offset(0, 16).Value = cl.Offset(0, 60).Value ' cancel date
                            ws2.Cells(currentRow, sourceCol).Offset(0, 17).Value = cl.Offset(0, 61).Value ' first PO#
                            ws2.Cells(currentRow, sourceCol).Offset(0, 18).Value = cl.Offset(0, 62).Value ' comments
                            
                            Exit For
                            
                        End If
                        
                        Next currentRow
                        
                    End If
        
                Next cl
        
            'hide the sheet "ItemInfo"
            curWeekUploadFile.Worksheets("ItemInfo").Visible = xlSheetHidden
            
            Application.DisplayAlerts = False
        
            curWeekUploadFile.Close SaveChanges:=True
            
            Application.DisplayAlerts = False
            
            Application.Wait (Now + TimeValue("00:00:02"))
            
            Next myItem
            
'**************************************************************************************************************
'OVERALL LOOP PER VENDOR FINISH
'**************************************************************************************************************

Application.DisplayAlerts = False

Workbooks("SKU_working_file.xlsx").Close SaveChanges:=False

Application.DisplayAlerts = True

Application.Wait (Now + TimeValue("00:00:02"))
    
Application.ScreenUpdating = True

Windows("create_SKU_import_files.xlsb").Activate

MsgBox "The files are ready for upload.", vbInformation, "Export Successful!"

End If

    End With
        
    End With
    
''turn on OneDrive sync
'ManageOnedriveSync (0)

'FOR THE NEXT REVISION
'close the extra window OneDrive opens on restart
'CloseWindow ("C:\Users\" & strUserName & "\OneDrive - COMPANY")

Exit Sub

NothingToSeeHere:

    Application.DisplayAlerts = False
                
    Workbooks("SKU_working_file.xlsx").Close SaveChanges:=False
    
    Application.DisplayAlerts = True
    
    MsgBox ("You need the file 'CreatedSKUsUploadPricing.xlsx' in the same directory as this workbook in order to proceed. This script will now exit."), vbCritical, "File Missing!"
    
'    'turn on OneDrive sync
'    ManageOnedriveSync (0)
    
    Exit Sub
    
NothingToSeeHere2:

    Application.DisplayAlerts = False
                
    Workbooks("SKU_working_file.xlsx").Close SaveChanges:=False
    
    Application.DisplayAlerts = True
    
    MsgBox ("You need the file 'CreatedSKUs2PA.xlsx' in the same directory as this workbook in order to proceed. This script will now exit."), vbCritical, "File Missing!"
                
'    'turn on OneDrive sync
'    ManageOnedriveSync (0)
    
    Exit Sub
    
''turn on OneDrive sync
'ManageOnedriveSync (0)
    
End Sub
'Sub CloseWindow(ByVal FullPathName As String)
'
'    Dim sh As Object
'    Dim w As Variant
'
'
'    Set sh = CreateObject("shell.application")
'    For Each w In sh.Windows
'        Debug.Print w.document.focuseditem.path
'        If w.document.focuseditem.path = FullPathName Then
'            w.Quit
'        End If
'    Next w
'
'End Sub
Sub DeleteUnwantedEntries(MyFileName As String, MyCellRange As String)

    Dim ws2 As Worksheet
    Dim curWeekUploadFile As Workbook
    Dim lRw3 As Long
    Dim j As Long
    Dim cl As Range
    
    Workbooks.Open Filename:=MyFileName
    
    Set curWeekUploadFile = ActiveWorkbook
    
    Set ws2 = curWeekUploadFile.Worksheets("ItemInfo")
    
    'make sure there is no data in the ItemInfo table
    
    lRw3 = ws2.Range(MyCellRange).SpecialCells(xlCellTypeLastCell).Row
    
    For j = lRw3 To 2 Step -1
    
        Set cl = Range("A" & j)

        If Application.CountA(cl.EntireRow) > 0 Then

            cl.EntireRow.Delete
            
        End If

    Next j
    
    Application.DisplayAlerts = False
    
    curWeekUploadFile.Save
    
    curWeekUploadFile.Close
    
    Application.DisplayAlerts = True

End Sub

'Check if the workbook is already open; if yes, do nothing, if not, open it

Function IsWorkBookOpen(Name As String) As Boolean

    Dim wb As Workbook
    On Error Resume Next
    Set wb = Application.Workbooks.item(Name)
    IsWorkBookOpen = (Not wb Is Nothing)
    
End Function

'eliminate duplicate values from the array with the vendor IDs

Function ArrayRemoveDups(MyArray As Variant) As Variant
    Dim nFirst As Long, nLast As Long, i As Long
    Dim item As String
    
    Dim arrTemp() As String
    Dim Coll As New Collection
 
    'Get First and Last Array Positions
    nFirst = LBound(MyArray) ' - 1 'deduct one to get rid of the empty element
    nLast = UBound(MyArray)
    ReDim arrTemp(nFirst To nLast)
 
    'Convert Array to String
    For i = nFirst To nLast
        arrTemp(i) = CStr(MyArray(i))
    Next i
    
    'Populate Temporary Collection
    On Error Resume Next
    For i = nFirst To nLast
        Coll.Add arrTemp(i), arrTemp(i)
    Next i
    Err.Clear
    On Error GoTo 0
 
    'Resize Array
    nLast = Coll.Count + nFirst - 1
    ReDim arrTemp(nFirst To nLast)
    
    'Populate Array
    For i = nFirst To nLast
        arrTemp(i) = Coll(i - nFirst + 1)
    Next i
    
    'Output Array
    ArrayRemoveDups = arrTemp
 
End Function

'switch sync to onedrive on or off

Sub ManageOnedriveSync(ByVal action As Integer)

    Dim shell As Object
    Set shell = VBA.CreateObject("WScript.Shell")
    Dim waitTillComplete As Boolean: waitTillComplete = False
    Dim style As Integer: style = 1
    Dim errorcode As Integer
    Dim path As String
    
    Dim commandAction As String
    Select Case action
    Case 1
        commandAction = "/shutdown"
    End Select

    path = Chr(34) & "%programfiles%\Microsoft OneDrive\Onedrive.exe" & Chr(34) & " " & commandAction

    errorcode = shell.Run(path, style, waitTillComplete)

End Sub

'check if the vendor ID array is empty

Function IsArrayAllocated(Arr As Variant) As Boolean

    On Error Resume Next
    IsArrayAllocated = IsArray(Arr) And _
                       Not IsError(LBound(Arr, 1)) And _
                       LBound(Arr, 1) <= UBound(Arr, 1)
End Function

