Attribute VB_Name = "AddInverseFilter"
Option Explicit

Public Sub AddToCellMenu()

Dim FilterMenu As CommandBarControl

    ' Delete the controls first to avoid duplicates
    Call DeleteFromCellMenu

    ' Set ContextMenu to the Cell context menu
    ' 31402 is the filter sub-menu of the cell context menu
    Set FilterMenu = Application.CommandBars("Cell").FindControl(ID:=31402)

    ' Add one custom button to the Cell context menu
    With FilterMenu.Controls.Add(Type:=msoControlButton, before:=3)
        .OnAction = "'" & ThisWorkbook.Name & "'!" & "InvertFilter"
        .FaceId = 1807
        .Caption = "Invert Filter Selection"
        .Tag = "My_Cell_Control_Tag"
    End With

End Sub

Private Sub DeleteFromCellMenu()

Dim FilterMenu As CommandBarControl
Dim ctrl As CommandBarControl

    ' Set ContextMenu to the Cell context menu
    ' 31402 is the filter sub-menu of the cell context menu
    Set FilterMenu = Application.CommandBars("Cell").FindControl(ID:=31402)

    ' Delete the custom controls with the Tag : My_Cell_Control_Tag
    For Each ctrl In FilterMenu.Controls
        If ctrl.Tag = "My_Cell_Control_Tag" Then
            ctrl.Delete
        End If
    Next ctrl

End Sub

Public Sub InvertFilter()

Application.ScreenUpdating = False

Dim cell As Range
Dim af As AutoFilter
Dim f As Filter
Dim i As Integer

Dim arrCur As Variant
Dim arrNew As Variant
Dim rngCol As Range
Dim c As Range
Dim txt As String
Dim bBlank As Boolean

    ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ' INITAL CHECKS
    ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    Set cell = ActiveCell

    If cell.Parent.AutoFilterMode = False Then
        MsgBox "No filters on current sheet"
        Exit Sub
    End If

    Set af = cell.Parent.AutoFilter

    If Application.Intersect(cell, af.Range) Is Nothing Then
        MsgBox "Current cell not part of filter range"
        Exit Sub
    End If

    i = cell.Column - af.Range.Cells(1, 1).Column + 1
    Set f = af.Filters(i)

    If f.On = False Then
        MsgBox "Current column not being filtered. Nothing to invert"
        Exit Sub
    End If

    ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ' GET CURRENT FILTER DATA
    ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    ' Single value criteria
    If f.Operator = 0 Then
        If f.Criteria1 = "<>" Then ArrayAdd arrNew, "="
        If f.Criteria1 = "=" Then ArrayAdd arrNew, "<>"
        ArrayAdd arrCur, f.Criteria1
    ' Pair of values used as criteria
    ElseIf f.Operator = xlOr Then
        ArrayAdd arrCur, f.Criteria1
        ArrayAdd arrCur, f.Criteria2
    ' Multi list criteria
    ElseIf f.Operator = xlFilterValues Then
        arrCur = f.Criteria1
    Else
        MsgBox "Current filter is not selecting values. Cannot process inversion"
        Exit Sub
    End If

    ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ' COMPUTE INVERTED FILTER DATA
    ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    ' Only process if new list is empty
    ' Being non-empty implies we're just toggling blank state and new list is already determined for that
    If IsEmpty(arrNew) Then

        ' Get column of data, ignoring header row
        Set rngCol = af.Range.Resize(af.Range.Rows.Count - 1, 1).Offset(1, i - 1)
        bBlank = False

        For Each c In rngCol

            ' Ignore blanks for now; they get special processing at the end
            If c.Text <> "" Then

                ' If the cell text is in neither the current filter list ...
                txt = "=" & c.Text
                If Not ArrayContains(arrCur, txt) Then

                    ' ... nor the new proposed list then add it to the new proposed list
                    If Not ArrayContains(arrNew, txt) Then ArrayAdd arrNew, txt

                End If

            Else
                ' Record that we have blank cells
                bBlank = True
            End If

        Next c

        ' Process blank options
        ' If we're not currently selecting for blanks ...
        ' ... and there are blanks ...
        ' ... then filter for blanks in new selection
        If (Not arrCur(UBound(arrCur)) = "=" And bBlank) Then ArrayAdd arrNew, "="

    End If

    ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ' APPLY NEW FILTER
    ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    Select Case UBound(arrNew)
        Case 0:
            MsgBox "Didn't find any values to invert"
            Exit Sub
        Case 1:
            af.Range.AutoFilter _
                Field:=i, _
                Criteria1:=arrNew(1)
        Case 2:
            af.Range.AutoFilter _
                Field:=i, _
                Criteria1:=arrNew(1), _
                Criteria2:=arrNew(2), _
                Operator:=xlOr
        Case Else:
            af.Range.AutoFilter _
                Field:=i, _
                Criteria1:=arrNew, _
                Operator:=xlFilterValues
    End Select

Application.ScreenUpdating = True

End Sub

Private Sub ArrayAdd(ByRef a As Variant, item As Variant)

Dim i As Integer

    If IsEmpty(a) Then
        i = 1
        ReDim a(1 To i)
    Else
        i = UBound(a) + 1
        ReDim Preserve a(1 To i)
    End If

    a(i) = item

End Sub

Private Function ArrayContains(a As Variant, item As Variant) As Boolean

Dim i As Integer

    If IsEmpty(a) Then
        ArrayContains = False
        Exit Function
    End If

    For i = LBound(a) To UBound(a)
        If a(i) = item Then
            ArrayContains = True
            Exit Function
        End If
    Next i

    ArrayContains = False

End Function

' Used to find the menu IDs
Private Sub ListMenuInfo()

Dim row As Integer
Dim Menu As CommandBarControl
Dim MenuItem As CommandBarControl
Dim SubMenuItem As CommandBarControl

    row = 1
    On Error Resume Next
    For Each Menu In CommandBars("cell").Controls
        For Each MenuItem In Menu.Controls
            For Each SubMenuItem In MenuItem.Controls
                Cells(row, 1) = Menu.Caption
                Cells(row, 2) = Menu.ID
                Cells(row, 3) = MenuItem.Caption
                Cells(row, 4) = MenuItem.ID
                Cells(row, 5) = SubMenuItem.Caption
                Cells(row, 6) = SubMenuItem.ID
                row = row + 1
            Next SubMenuItem
        Next MenuItem
    Next Menu

End Sub

