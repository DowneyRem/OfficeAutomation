Attribute VB_Name = "模块1"
Sub Auto_Close()
    
    
    File = ActiveWorkbook.Path & "/" & ActiveWorkbook.Name
    If InStr(1, File, "sharepoint.com", vbBinaryCompare) > 0 Then
        '获取OneDrive在线链接，采用mid截取掉前段云端路径
        'Substitute(ActiveWorkbook.Path, "/", "@", 4) 第四参数，需要根据实际情况修改
        Path = Mid(ActiveWorkbook.Path, Application.Find("@", Application.Substitute(ActiveWorkbook.Path, "/", "@", 6)), 999)
        File = Environ("OneDrive") & Path & "/" & ActiveWorkbook.Name
    End If
    
    File = Replace(File, "xlsm", "xlsx")
    MsgBox (File)
    
    If InStr(1, ActiveWorkbook.Name, "唐门小说更新统计.xlsm", vbBinaryCompare) > 0 Then
        '选中单元格，输入上次保存时间，保存
        lastsave = ActiveWorkbook.BuiltinDocumentProperties("Last Save Time")
        lastsave = VBA.Int(lastsave)
        Sheets("小说").Range("J1").Value = "更新时间: " & Application.Text(lastsave, "yyyy/mm/dd")
        
        '使用SaveCopyAs 无需设置文件格式
        ActiveWorkbook.Save
        ActiveWorkbook.SaveCopyAs (File)
        
    Else
        MsgBox ("文件错误！" & vbCrLf & "当前文件：" & ActiveWorkbook.Name)
    End If
    
    
End Sub







