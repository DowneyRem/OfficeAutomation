## Excel2

### 功能需求分析

#### Python 处理部分
0. pywin32 打开Excel
1. 在表格内填入表格上次保存时间
2. 关闭时，自动备份（将xlsm另存为xlsx）

#### Excel VBA 处理部分
0. （手动）打开Excel
1. 在表格内填入表格上次保存时间
2. 关闭时，自动备份（将xlsm另存为xlsx）

### 实现过程

#### Python 实现过程
0. pywin32 打开Excel
1. 获取上次文件修改时间
2. saveascopy 备份工作簿，将xlsm另存为xlsx
2. ~~使用 WPS Office，分享 xlsx文件~~

#### Excel VBA 实现过程
0. 打开Excel，关闭时自动运行 aoto_close 宏代码
1. 获取上次文件修改时间
2. saveascopy 备份工作簿，将xlsm另存为xlsx
3. ~~使用 WPS Office，分享 xlsx文件~~

[更多内容](https://kdocs.cn/l/cgEZR9rebTTE)


## 自动宏
### [Word 自动宏](https://docs.microsoft.com/zh-cn/office/vba/word/concepts/customizing-word/auto-macros)

| 宏名      | 运行条件                   |
| --------- | -------------------------- |
| AutoExec  | 启动 Word 或加载全局模板时 |
| AutoNew   | 每次新建文档时             |
| AutoOpen  | 每次打开已有文档时         |
| AutoClose | 每次关闭文档时             |
| AutoExit  | 退出 Word 或卸载共用模板时 |

> AutoExec 宏，只有存储于以下位置时才可自动运行：Normal 模板、通过 **“模板和加载项”** 对话框全局加载的模板、或由“Startup”文件夹指定的文件夹中的全局模板。 
>
> 在命名冲突的情况下（多个自动宏名相同），Word 将运行最近的上下文中的自动宏。 例如，如果同时在文档及其附加的模板中创建了 AutoClose 宏，则仅执行文档中的自动宏。 
>
> 如果在 Normal 模板中创建了 AutoNew 宏，只有当文档或其附加的模板中没有名为 AutoNew 的宏时，该自动宏才能运行。

### [Excel 自动宏](https://docs.microsoft.com/zh-cn/office/vba/api/excel.xlrunautomacro)

| 宏名                 |                    |
| :------------------- | ------------------ |
| **xlAutoActivate**   | Auto_Activate 宏   |
| **xlAutoClose**      | Auto_Close 宏      |
| **xlAutoDeactivate** | Auto_Deactivate 宏 |
| **xlAutoOpen**       | Auto_Open 宏       |

> 运行附属于指定工作簿的 Auto_Open、Auto_Close、Auto_Activate 或 Auto_Deactivate 宏。 保留本方法是为了保持向后兼容性。
>
> 对于新的Visual Basic代码，应该使用 **Open**、**Activate** 和 **Deactivate** 事件以及 **Close** 方法，而不是这些宏。
>
> https://docs.microsoft.com/zh-cn/office/vba/api/excel.workbook.runautomacros
