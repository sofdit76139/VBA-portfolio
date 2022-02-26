Attribute VB_Name = "mod_AddLeadingZeros"
Sub AddLeadingZerosToNumber()

Dim ws As Worksheet
Dim wb As Workbook
Dim fileName As String
Dim wsName As String
Dim wsRange As String
Dim cl As Range

    fileName = Application.InputBox(prompt:="Please enter the name of the workbook (Example: example.xlsx):", Type:=2)
    
    If fileName = "" Then
    
        MsgBox "Input is neded to proceed.", vbCritical, "Information missing!"
        Exit Sub
        
    End If
    
    wsName = Application.InputBox(prompt:="Please enter the name of the worksheet:", Type:=2)
    
    If wsName = "" Then
    
        MsgBox "Input is neded to proceed.", vbCritical, "Information missing!"
        Exit Sub
        
    End If
    
    wsRange = Application.InputBox(prompt:="Please enter the range where you want to add the leading zeros (Example: A1:A15):", Type:=2)
    
    If wsRange = "" Then
    
        MsgBox "Input is neded to proceed.", vbCritical, "Information missing!"
        Exit Sub
        
    End If
    
    Set wb = Workbooks(fileName)
    
    Set ws = wb.Worksheets(wsName)
    
    IsWorkBookOpen (fileName)
    
    For Each cl In ws.Range(wsRange)
    
        If Len(cl.Value) < 12 Then
        
            cl.NumberFormat = "@"
            
            cl.Value = "0" & cl.Value
        
        End If
        
    Next cl

    MsgBox "Leading zeros were added.", vbInformation, "Done"

End Sub

'Check if the workbook is already open; if yes, do nothing, if not, open it

Function IsWorkBookOpen(Name As String) As Boolean

    Dim wb As Workbook
    On Error Resume Next
    Set wb = Application.Workbooks.item(Name)
    IsWorkBookOpen = (Not wb Is Nothing)
    
End Function
