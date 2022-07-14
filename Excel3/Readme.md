## Excel3

### 具体需求：
- 按人名拆分表格，并打印表头，如下图所示

- 原始数据

- 要求效果


### 实现思路：
1. 表头部分通过【打印标题】重复打印
2. 数据部分通过【插分页符】分页打印
3. 通过【VBA】循环判断来插入分页符

### 解决方法：
0. 切换至打印视图，分页预览
1. 设置标题打印
2. VBA 代码插入分页符
3. 批量打印

```  VBA
Sub 插入分页符()
'非空行前，插入分页符
'
    On Error Resume Next:
    For i = 2 To 1160
        Text = Range("A" & i).Value
        If Text <> "" Then
            Debug.Print Text
            Range("A" & i & ":B" & i).Select
            ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
        End If
    Next
End Sub
```

### 实现效果：
- 按人名拆分表格，并打印表头
