Attribute VB_Name = "模块1"
Sub 自动1()
Attribute 自动1.VB_ProcData.VB_Invoke_Func = " \n14"


    If InStr(1, ActiveSheet.Name, "赞", vbTextCompare) > 0 Or InStr(1, ActiveSheet.Name, "点", vbTextCompare) > 0 Then
    
        '选中末行末列的单元格
        ActiveCell.SpecialCells(xlLastCell).Select
        
        '左移1列，插入1列
        ActiveCell.Offset(0, -1).Columns("A:A").EntireColumn.Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        
        '右移2列，复制数据
        ActiveCell.Offset(0, 2).Columns("A:A").EntireColumn.Select
        Selection.Copy
        
        '左移2列，粘贴数据
        ActiveCell.Offset(0, -2).Columns("A:A").EntireColumn.Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
        
        '数据首行，加入日期
        Selection.End(xlUp).Select
        ActiveCell.FormulaR1C1 = Date
        
        '点赞量数据表中，修改表头数据
        If InStr(1, ActiveSheet.Name, "赞", vbTextCompare) > 0 Then
            ActiveCell.Offset(0, -2).Range("A1:B1").Select
            Selection.AutoFill Destination:=ActiveCell.Range("A1:E1"), Type:=xlFillDefault
            ActiveCell.Offset(0, 2).Columns("A:A").EntireColumn.Select
        End If
        
        '左移4列，隐藏该列
        ActiveCell.Offset(0, -4).Columns("A:A").EntireColumn.Select
        Selection.EntireColumn.Hidden = True
        
        
        '保存
        ActiveWorkbook.Save
        
    End If
    
End Sub


Sub 自动2()


    If InStr(1, ActiveSheet.Name, "个人", vbTextCompare) > 0 Then
    
        'Sheets("个人").Select
        Range("A1:H30").Select
        Selection.ClearContents
        
        Sheets("累计").Visible = True
        Sheets("累计").Select
        '筛选数据
        ActiveSheet.Range("$A$1:$H$51").AutoFilter Field:=2, Criteria1:="文"
        ActiveSheet.Range("$A$1:$H$51").AutoFilter Field:=3, Criteria1:="否"
        
        
        '复制，粘贴
        Range("A1:H43").Select
        Selection.Copy
        Sheets("个人").Select
        Range("A1").Select
        ActiveSheet.Paste
        Selection.AutoFilter
        
        ActiveWorkbook.Worksheets("个人").AutoFilter.Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("个人").AutoFilter.Sort.SortFields.Add2 Key:=Range( _
            "H1:H30"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        With ActiveWorkbook.Worksheets("个人").AutoFilter.Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
        '隐藏表单
        Sheets("累计").Visible = False
    End If
    
End Sub


Sub 自动3()


    If InStr(1, ActiveSheet.Name, "近期", vbTextCompare) > 0 Then
        
        Range("G2").Select
        ActiveWorkbook.Worksheets("近期").AutoFilter.Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("近期").AutoFilter.Sort.SortFields.Add2 Key:=Range( _
            "G2:G44"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
            
        With ActiveWorkbook.Worksheets("近期").AutoFilter.Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
        Range("K2").Select
        ActiveWorkbook.Worksheets("近期").AutoFilter.Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("近期").AutoFilter.Sort.SortFields.Add2 Key:=Range( _
            "K2:K44"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
            
        With ActiveWorkbook.Worksheets("近期").AutoFilter.Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
    End If
    
End Sub


Sub 自动化()
Attribute 自动化.VB_ProcData.VB_Invoke_Func = " \n14"


    '选中工作表，运行宏
    Sheets("点").Visible = True
    Sheets("点").Select
    Application.Run "唐门小说点赞统计.xlsm!自动1"
    
    Sheets("赞").Visible = True
    Sheets("赞").Select
    Application.Run "唐门小说点赞统计.xlsm!自动1"
    
    Sheets("个人").Select
    Application.Run "唐门小说点赞统计.xlsm!自动2"
        
    Sheets("近期").Select
    Application.Run "唐门小说点赞统计.xlsm!自动3"
    
    Sheets("点").Visible = False
    Sheets("赞").Visible = False
    Sheets("名称").Visible = False
    ActiveWorkbook.Save
    
    
End Sub
