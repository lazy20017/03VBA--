' ========================================
' Login Module - Handle user login verification
' ========================================

' 文档打开时自动运行
Sub Document_Open()
    On Error Resume Next
    ShowLoginForm
    If Err.Number <> 0 Then
        MsgBox "启动登录失败: " & Err.Description, vbCritical, "错误"
    End If
    On Error GoTo 0
End Sub

' 简单的登录运行过程
Sub RunLogin()
    ShowLoginForm
End Sub

' 显示登录窗体（使用 InputBox）
Sub ShowLoginForm()
    Dim StudentID As String
    Dim Name As String
    
    ' 输入学号
    StudentID = InputBox("请输入学号:", "学生登录")
    If StudentID = "" Then
        MsgBox "学号不能为空！", vbExclamation, "提示"
        Exit Sub
    End If
    
    ' 输入姓名
    Name = InputBox("请输入姓名:", "学生登录")
    If Name = "" Then
        MsgBox "姓名不能为空！", vbExclamation, "提示"
        Exit Sub
    End If
    
    ' 验证登录
    If VerifyLogin(StudentID, Name) Then
        RecordLoginInfo StudentID, Name
        MsgBox "登录成功！欢迎 " & Name & "！", vbInformation, "欢迎"
    End If
End Sub

' 验证登录
Function VerifyLogin(StudentID As String, Name As String) As Boolean
    ' 这里可以添加实际的验证逻辑
    ' 暂时允许任意输入
    If Len(StudentID) > 0 And Len(Name) > 0 Then
        VerifyLogin = True
    Else
        VerifyLogin = False
        MsgBox "学号或姓名无效！", vbCritical, "登录失败"
    End If
End Function

' 记录登录信息
Sub RecordLoginInfo(StudentID As String, Name As String)
    ' 在文档中记录登录信息
    Dim content As String
    content = "学号: " & StudentID & ", 姓名: " & Name & ", 登录时间: " & Now
    
    ' 添加到文档末尾
    ThisDocument.Content.InsertAfter vbCrLf & content
    
    Debug.Print "登录信息: " & content
End Sub
