Attribute VB_Name = "Mod_LooksLikeChristmas"
Option Explicit

Sub LooksLikeChristmas()
' Determine if a product is in one of the Christmas/Holiday categories/subcatgories

Dim MyArray() As Variant
Dim DataRange As Range
Dim DataRow As Range
Dim DataCell As Range
Dim MyRowNumber As Long
Dim IsInArray As Boolean

Dim CatSearchString As String
Dim i As Long
Dim lRw As Long

Application.ScreenUpdating = False

On Error Resume Next

'remove all filters
ActiveSheet.ShowAllData

'declare the array; all categories are Christmas-related
    MyArray() = Array("661", "662", "663", "664", "665", "666", "667", "668", "669", "670", "671", "672", "673", "681", "695", "696", "807", "808", "809", "810", "811", "812", "813", "816", "861", "864", "866", "903", "919")

'Determine the active range's last row

    lRw = ActiveSheet.Range("A:A").SpecialCells(xlCellTypeLastCell).Row
    
    Set DataRange = ActiveSheet.Range("R2:AB" & lRw)
    
'Determine the search value, then see if it is contained in the array, stop when the first value is found

For Each DataRow In DataRange.Rows

    For Each DataCell In DataRow.Cells

        CatSearchString = DataCell.Value
        
        MyRowNumber = DataCell.Row
        
        If CatSearchString = "" Then
        
            CatSearchString = "000"
            
        End If
        
        IsInArray = (UBound(Filter(MyArray, CatSearchString)) > -1)
        
        If IsInArray = True Then
        
            Cells(MyRowNumber, 30).Value = "1"
            Cells(MyRowNumber, 31).Value = CatSearchString
            
            Exit For
        Else
        
            Cells(MyRowNumber, 30).Value = "0"
            
        End If
        
    Next DataCell

Next DataRow

Application.ScreenUpdating = True
    
End Sub
