
' Remove Duplicates - Single column
mySheet.Range("P:P").RemoveDuplicates Columns:=1, Header:=xlYes


' Remove Duplicates - 2 columns
mySheet.Range("M:N").RemoveDuplicates Columns:=Array(1, 2), Header:=xlYes

' Merge
With mySheet.Range(scol & "4:" & ecol & "4")
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Merge
End With


' Sorting
mySheet.Sort.SortFields.Clear
mySheet.Sort.SortFields.Add2 Key:=Range("M1:M" & Rows.Count), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
mySheet.Sort.SortFields.Add2 Key:=Range("N1:N" & Rows.Count), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With mySheet.Sort
    .SetRange Range("M1:N" & Rows.Count)
    .Header = xlGuess
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With


'Format Borders
With mySheet.Range("A5:" & planned_APSW_end_col & LastRow).Borders
    .LineStyle = 1 'xlContinuous
    .ColorIndex = 15
    .Weight = xlThin
End With


' Auto Filter
If Not mySheet.AutoFilterMode Then
    mySheet.Rows("5:5").AutoFilter
Else
    mySheet.AutoFilter.ShowAllData
    mySheet.Rows("5:5").AutoFilter
End If

