Attribute VB_Name = "ConvertSKUtoHTML"
Sub ConvertSKUtoHTML()

Dim ws As Worksheet

Dim wb As Workbook

Dim xRet As Boolean

Application.ScreenUpdating = False

On Error Resume Next


xRet = IsWorkBookOpen("2021 Black Friday.xlsx") 'INCLUDE FULL NAME OF WORKBOOK HERE

If xRet Then
    'MsgBox "The file is open."
Else
    Workbooks.Open Filename:="C:\Users\Sofie.Dittmann\OneDrive - COMPANY\Reporting\Merchandising\2021 Black Friday.xlsx" 'INCLUDE FULL PATH TO WORKBOOK HERE
End If

Set wb = Application.Workbooks("2021 Black Friday.xlsx") 'INCLUDE FULL NAME OF WORKBOOK HERE

Set ws = wb.Worksheets("BlackFriday") 'INSERT SHEET NAME HERE

Dim lRw As Long

lRw = ActiveSheet.Range("A:A").SpecialCells(xlCellTypeLastCell).row - 1

Dim i As Long

With ws
    
    For i = 1 To lRw Step 1
        
        .Cells(i + 2, 1).NumberFormat = "@" 'COLUMN A
        
        .Cells(i + 1, 1).Activate
        
'        .Cells(i + 2, 2).NumberFormat = "@" 'COLUMN B
'
'        .Cells(i + 1, 2).Activate
'        .Hyperlinks.Add Anchor:=ActiveCell, Address:=ActiveCell.Text, TextToDisplay:=ActiveCell.Text ' turn all items into hyperlinks using the URLs already there
        .Hyperlinks.Add Anchor:=ActiveCell, Address:="https://COMPANY.com/search?w=" & ActiveCell.Text, TextToDisplay:=ActiveCell.Text ' turn all SKUs etc. into hyperlinks
'        .Hyperlinks.Add Anchor:=ActiveCell, Address:="https://www.google.com/search?q=" & ActiveCell.Text, TextToDisplay:=ActiveCell.Text ' turn all SKUs etc. into hyperlinks

    Next i
    
End With

Application.ScreenUpdating = True

MsgBox "The HTML conversion is now complete."

End Sub

'Check if the workbook is already open; if yes, do nothing, if not, open it

Function IsWorkBookOpen(Name As String) As Boolean

    Dim wb As Workbook
    On Error Resume Next
    Set wb = Application.Workbooks.item(Name)
    IsWorkBookOpen = (Not wb Is Nothing)
    
End Function

