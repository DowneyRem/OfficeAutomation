## Excel4

### 具体需求：
- 按人名拆分表格，并打印表头，如下图所示
- 原始数据

![原始数据.png](https://raw.githubusercontent.com/DowneyRem/OfficeAutomation/main/Excel4/原始数据.png)
- 要求效果
每人1张对应的数据表格

![要求效果.png](https://raw.githubusercontent.com/DowneyRem/OfficeAutomation/main/Excel4/要求效果.png)

### 实现思路：
1. 表头部分通过【打印标题】重复打印
2. 数据部分通过【插分页符】分页打印
3. 通过【VBA】循环判断来插入分页符
4. 批量打印

### 解决方法：

0. 切换视图至，分页预览

![视图-分页预览.png](https://raw.githubusercontent.com/DowneyRem/OfficeAutomation/main/Excel4/视图-分页预览.png)

1. 设置标题打印

![页面布局-打印标题.png](https://raw.githubusercontent.com/DowneyRem/OfficeAutomation/main/Excel4/页面布局-打印标题.png)

![打印标题.png](https://raw.githubusercontent.com/DowneyRem/OfficeAutomation/main/Excel4/打印标题.png)

2. VBA 代码插入分页符
	- 录制宏：单独插入分页符
	
	![开发工具-录制宏.png](https://raw.githubusercontent.com/DowneyRem/OfficeAutomation/main/Excel4/开发工具-录制宏.png)
	
	![页面布局-插入分页符.png](https://raw.githubusercontent.com/DowneyRem/OfficeAutomation/main/Excel4/页面布局-插入分页符.png)
	
	- 修改宏：批量插入分页符
	  根据具体需求进行修改
	
```  VBA
Sub 插入分页符()
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

- 插入完成，效果如下：

![插入分页符效果.png](https://raw.githubusercontent.com/DowneyRem/OfficeAutomation/main/Excel4/插入分页符效果.png)

3. 批量打印

![批量打印.png](https://raw.githubusercontent.com/DowneyRem/OfficeAutomation/main/Excel4/批量打印.png)


### 实现效果：
- 按人名拆分表格，并打印表头

![要求效果.png](https://raw.githubusercontent.com/DowneyRem/OfficeAutomation/main/Excel4/要求效果.png)