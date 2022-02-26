Attribute VB_Name = "mod_CreateKeywords"
Sub MakeKeywordVariatons()

'Application.ScreenUpdating = False
'
'On Error Resume Next

    Dim rngPrefixesList, rngKeywordList As Range
    Dim rngPrefix, rngKeyword As Range
    Dim strVariationList As String
    
    Set rngPrefixesList = Sheet10.Range(Sheet10.Range("A2"), Sheet10.Range("A2").End(xlDown))
    Set rngKeywordList = Sheet10.Range(Sheet10.Range("B2"), Sheet10.Range("B2").End(xlDown))

    Sheet10.Range("D2").Select

    For Each rngPrefix In rngPrefixesList
        For Each rngKeyword In rngKeywordList
            ActiveCell = rngPrefix.Value & " " & rngKeyword
            If strVariationList = "" Then
                strVariationList = ActiveCell
            Else
                strVariationList = strVariationList & ", " & ActiveCell
            End If
            ActiveCell.Offset(1, 0).Select
            
        Next
    Next
    
    Sheet10.Range("G2") = strVariationList
    
'    Sheet4.Range("K2") = strVariationList
    
Application.ScreenUpdating = True

End Sub
