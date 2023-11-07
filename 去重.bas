Attribute VB_Name = "模块1"
Sub quchong()
    '功能去除A列中重复值，结果放入B列
    Dim data_dic As Object
    Set data_dic = CreateObject("scripting.dictionary")
    r_max = Range("a1048576").End(xlUp).Row
    For i = 1 To r_max
'        On Error Resume Next
        data_dic(Cells(i, 1).Value) = i
    Next i
    Range("B1").Resize(data_dic.Count) = Application.Transpose(data_dic.keys)
    Set data_dic = Nothing
    
    
    Columns("B:B").Select
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add Key:=Range("b1:b838") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range("b1:b838")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

