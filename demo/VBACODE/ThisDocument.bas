' ========================================
' ThisDocument - Document-level events
' ========================================

' 文档打开时自动运行
Private Sub Document_Open()
    On Error Resume Next
    RunLogin
    If Err.Number <> 0 Then
        Debug.Print "自动运行错误: " & Err.Description
    End If
    On Error GoTo 0
End Sub

' 文档打开时自动运行 (备用方法)
Private Sub AutoOpen()
    On Error Resume Next
    RunLogin
    If Err.Number <> 0 Then
        Debug.Print "自动运行错误: " & Err.Description
    End If
    On Error GoTo 0
End Sub
