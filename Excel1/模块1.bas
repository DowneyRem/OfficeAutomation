Attribute VB_Name = "ģ��1"
Sub �Զ�1()
Attribute �Զ�1.VB_ProcData.VB_Invoke_Func = " \n14"


    If InStr(1, ActiveSheet.Name, "��", vbTextCompare) > 0 Or InStr(1, ActiveSheet.Name, "��", vbTextCompare) > 0 Then
    
        'ѡ��ĩ��ĩ�еĵ�Ԫ��
        ActiveCell.SpecialCells(xlLastCell).Select
        
        '����1�У�����1��
        ActiveCell.Offset(0, -1).Columns("A:A").EntireColumn.Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        
        '����2�У���������
        ActiveCell.Offset(0, 2).Columns("A:A").EntireColumn.Select
        Selection.Copy
        
        '����2�У�ճ������
        ActiveCell.Offset(0, -2).Columns("A:A").EntireColumn.Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
        
        '�������У���������
        Selection.End(xlUp).Select
        ActiveCell.FormulaR1C1 = Date
        
        '���������ݱ��У��޸ı�ͷ����
        If InStr(1, ActiveSheet.Name, "��", vbTextCompare) > 0 Then
            ActiveCell.Offset(0, -2).Range("A1:B1").Select
            Selection.AutoFill Destination:=ActiveCell.Range("A1:E1"), Type:=xlFillDefault
            ActiveCell.Offset(0, 2).Columns("A:A").EntireColumn.Select
        End If
        
        '����4�У����ظ���
        ActiveCell.Offset(0, -4).Columns("A:A").EntireColumn.Select
        Selection.EntireColumn.Hidden = True
        
        
        '����
        ActiveWorkbook.Save
        
    End If
    
End Sub


Sub �Զ�2()


    If InStr(1, ActiveSheet.Name, "����", vbTextCompare) > 0 Then
    
        'Sheets("����").Select
        Range("A1:H30").Select
        Selection.ClearContents
        
        Sheets("�ۼ�").Visible = True
        Sheets("�ۼ�").Select
        'ɸѡ����
        ActiveSheet.Range("$A$1:$H$51").AutoFilter Field:=2, Criteria1:="��"
        ActiveSheet.Range("$A$1:$H$51").AutoFilter Field:=3, Criteria1:="��"
        
        
        '���ƣ�ճ��
        Range("A1:H43").Select
        Selection.Copy
        Sheets("����").Select
        Range("A1").Select
        ActiveSheet.Paste
        Selection.AutoFilter
        
        ActiveWorkbook.Worksheets("����").AutoFilter.Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("����").AutoFilter.Sort.SortFields.Add2 Key:=Range( _
            "H1:H30"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        With ActiveWorkbook.Worksheets("����").AutoFilter.Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
        '���ر�
        Sheets("�ۼ�").Visible = False
    End If
    
End Sub


Sub �Զ�3()


    If InStr(1, ActiveSheet.Name, "����", vbTextCompare) > 0 Then
        
        Range("G2").Select
        ActiveWorkbook.Worksheets("����").AutoFilter.Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("����").AutoFilter.Sort.SortFields.Add2 Key:=Range( _
            "G2:G44"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
            
        With ActiveWorkbook.Worksheets("����").AutoFilter.Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
        Range("K2").Select
        ActiveWorkbook.Worksheets("����").AutoFilter.Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("����").AutoFilter.Sort.SortFields.Add2 Key:=Range( _
            "K2:K44"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
            
        With ActiveWorkbook.Worksheets("����").AutoFilter.Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
    End If
    
End Sub


Sub �Զ���()
Attribute �Զ���.VB_ProcData.VB_Invoke_Func = " \n14"


    'ѡ�й��������к�
    Sheets("��").Visible = True
    Sheets("��").Select
    Application.Run "����С˵����ͳ��.xlsm!�Զ�1"
    
    Sheets("��").Visible = True
    Sheets("��").Select
    Application.Run "����С˵����ͳ��.xlsm!�Զ�1"
    
    Sheets("����").Select
    Application.Run "����С˵����ͳ��.xlsm!�Զ�2"
        
    Sheets("����").Select
    Application.Run "����С˵����ͳ��.xlsm!�Զ�3"
    
    Sheets("��").Visible = False
    Sheets("��").Visible = False
    Sheets("����").Visible = False
    ActiveWorkbook.Save
    
    
End Sub
