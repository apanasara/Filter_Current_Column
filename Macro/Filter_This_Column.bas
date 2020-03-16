'------------------------------------Filtering the data-------------------------------------
Sub Filter_contains()
    Dim FilterRange As Range
    Dim ColumnName As String, fltr As String
    Dim col As Integer
    
    Application.ScreenUpdating = False
    
    On Error Resume Next
    '----------Filters are there in sheet or not?-------------
    If Not (ActiveSheet.AutoFilterMode) And ActiveCell.ListObject Is Nothing Then
         MsgBox "Filter-Range is not applied, please apply filters"
         Exit Sub
    End If
    
    '-------What is to be filtered?--------------
    fltr = InputBox("Filtering value", ActiveSheet.Name & " : Filtering Active Column", ActiveCell.Value)
    If fltr = "" Or fltr = "+" Then: Exit Sub
    
    '-------In which column filter to be applied?--------------
    If Not (ActiveCell.ListObject Is Nothing) Then
        Set FilterRange = ActiveCell.ListObject.HeaderRowRange
    Else
        Set FilterRange = ActiveSheet.AutoFilter.Range
    End If
    
    ColumnName = Intersect(FilterRange.Rows(1), ActiveCell.EntireColumn).Value
    col = Application.Match(ColumnName, FilterRange.Rows(1), 0)

    '------------applying new filter / adding into existing filter?-------
    If Left(fltr, 1) <> "+" Then
         ActiveSheet.ShowAllData 'if data is not filltered then will show error
    Else
          fltr = Right(fltr, Len(fltr) - 1) 'removing + sign
    End If
    
    
    '--------Finally Applying filter--------
    If Not (ActiveCell.ListObject Is Nothing) Then
        ActiveCell.ListObject.Range.AutoFilter _
            Field:=col, _
            Criteria1:="=*" & fltr & "*", Operator:=xlOr, Criteria2:="=" & fltr
    Else
        ActiveSheet.Range(FilterRange.Address).AutoFilter _
            Field:=col, _
            Criteria1:="=*" & fltr & "*", Operator:=xlOr, Criteria2:="=" & fltr
    End If
    
    Application.ScreenUpdating = False
End Sub
