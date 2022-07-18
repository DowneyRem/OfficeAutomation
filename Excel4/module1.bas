Attribute VB_Name = "Macro1"
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
