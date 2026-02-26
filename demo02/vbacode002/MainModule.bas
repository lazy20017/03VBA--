' -*- coding: utf-8 -*-
' =====================================================
' 学生作业保护系统 - 主模块
' 功能：全局变量定义和系统入口
' =====================================================

Option Explicit

' Word 常量声明
Const wdAllowOnlyReading As Long = -1
Const wdAllowOnlyFormFields As Long = 3

' 超级管理员账号和密码
Const SUPER_ADMIN_ID As String = "admin"
Const SUPER_ADMIN_PWD As String = "admin123"

' 全局变量
Public g_StudentID As String          ' 当前登录学生学号
Public g_StudentName As String         ' 当前登录学生姓名
Public g_IsReturningStudent As Boolean ' 是否为Returning学生
Public g_IsLoggedIn As Boolean         ' 是否已登录
Public g_IsSuperAdmin As Boolean       ' 是否为超级管理员

' =====================================================
' 系统入口 - 文档打开时自动运行
' =====================================================

' 显示登录窗体
Public Sub ShowLoginForm()
    Dim studentID As String
    Dim studentName As String
    Dim isReturning As Boolean
    Dim isFirstLogin As Boolean
    Dim savedStudentID As String
    Dim savedStudentPwd As String
    Dim loginPrompt As String
    
    ' 检查是否已登录，防止重复运行
    If g_IsLoggedIn Then
        Exit Sub
    End If
    
    ' 先检查是否已有保存的学号（锁定检查）
    savedStudentID = GetSavedStudentID()
    savedStudentPwd = GetSavedStudentPwd()
    
    ' 根据是否首次登录显示不同的提示
    If savedStudentID = "" Then
        ' 首次登录
        loginPrompt = "请输入学号（首次登录将记录您的学号和姓名）:" & vbCrLf & vbCrLf & _
                     "提示：学号和姓名将成为登录凭证，请妥善保管！"
    Else
        ' 非首次登录
        loginPrompt = "请输入学号（该作业已锁定，请使用正确的学号和姓名登录）:"
    End If
    
    ' 输入账号（学号）
    On Error Resume Next
    studentID = Trim(InputBox(loginPrompt, "学生作业登录"))
    If studentID = "" Or Err.Number <> 0 Then
        MsgBox "账号不能为空！文档将关闭。", vbExclamation, "提示"
        ThisDocument.Close SaveChanges:=False
        Exit Sub
    End If
    Err.Clear
    
    ' 根据是否首次登录显示不同的密码提示
    If savedStudentID = "" Then
        loginPrompt = "请输入姓名（将作为登录密码）:" & vbCrLf & vbCrLf & _
                     "提示：姓名将成为登录密码，请正确输入！"
    Else
        loginPrompt = "请输入姓名（密码）:"
    End If
    
    ' 输入密码（姓名）
    studentName = Trim(InputBox(loginPrompt, "学生作业登录"))
    If studentName = "" Or Err.Number <> 0 Then
        MsgBox "密码不能为空！文档将关闭。", vbExclamation, "提示"
        ThisDocument.Close SaveChanges:=False
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 检查是否为超级管理员
    If studentID = SUPER_ADMIN_ID And studentName = SUPER_ADMIN_PWD Then
        ' 超级管理员：完全访问权限
        MsgBox "欢迎管理员！" & vbCrLf & vbCrLf & _
               "您拥有完全访问权限。", vbInformation, "管理员登录"
        
        ' 解除所有保护
        On Error Resume Next
        ThisDocument.Unprotect
        ThisDocument.Unprotect Password:="teacher2024"
        ThisDocument.Unprotect Password:="StudentReadOnly2024"
        Err.Clear
        On Error GoTo 0
        
        g_StudentID = studentID
        g_StudentName = studentName
        g_IsSuperAdmin = True
        g_IsLoggedIn = True
        
        ' 显示管理员功能菜单
        Call ShowAdminMenu
        Exit Sub
    End If
    
    ' 初始化登录状态
    isReturning = False
    isFirstLogin = False
    
    ' 检查学号是否已锁定（锁定逻辑）
    If savedStudentID <> "" Then
        ' 已有保存的学号，检查是否匹配
        If studentID <> savedStudentID Then
            ' 学号不匹配：关闭文档
            MsgBox "登录失败！" & vbCrLf & vbCrLf & _
                   "该作业已锁定，只能由授权用户打开。" & vbCrLf & _
                   "请使用正确的账号和密码登录。" & vbCrLf & vbCrLf & _
                   "【调试信息，请截图给老师】" & vbCrLf & _
                   "保存的账号(学号): '" & savedStudentID & "'" & vbCrLf & _
                   "保存的密码(姓名): '" & savedStudentPwd & "'" & vbCrLf & _
                   "本次输入的账号: '" & studentID & "'" & vbCrLf & _
                   "本次输入的密码(姓名): '" & studentName & "'", _
                   vbCritical, "登录失败"
            
            ThisDocument.Close SaveChanges:=False
            Exit Sub
        Else
            ' 学号匹配：检查密码（姓名）是否正确
            If studentName <> savedStudentPwd Then
                ' 密码不匹配：关闭文档
                MsgBox "登录失败！" & vbCrLf & vbCrLf & _
                       "密码错误，该作业已锁定。" & vbCrLf & _
                       "请使用正确的账号和密码登录。" & vbCrLf & vbCrLf & _
                       "【调试信息，请截图给老师】" & vbCrLf & _
                       "保存的账号(学号): '" & savedStudentID & "'" & vbCrLf & _
                       "保存的密码(姓名): '" & savedStudentPwd & "'" & vbCrLf & _
                       "本次输入的账号: '" & studentID & "'" & vbCrLf & _
                       "本次输入的密码(姓名): '" & studentName & "'", _
                       vbCritical, "登录失败"
                
                ThisDocument.Close SaveChanges:=False
                Exit Sub
            End If
        End If
        
        ' 学号和密码都匹配：后续登录
        isReturning = True
    Else
        ' 首次登录：保存学号和密码并允许编辑
        Call SaveStudentCredentials(studentID, studentName)
        
        ' 调试信息：显示即将保存的学号和密码
        MsgBox "【调试信息，请截图给老师】" & vbCrLf & _
               "第一次登录后，准备锁存的账号信息如下：" & vbCrLf & vbCrLf & _
               "即将保存的账号(学号)：'" & studentID & "'" & vbCrLf & _
               "即将保存的密码(姓名)：'" & studentName & "'", _
               vbInformation, "首次登录调试"
        
        isFirstLogin = True
        
        ' 确保文档保存（否则自定义属性可能丢失）
        ThisDocument.Save
    End If
    
    ' 初始化变量
    g_StudentID = studentID
    g_StudentName = studentName
    g_IsSuperAdmin = False
    
    ' 记录登录
    Call StudentStorage.RecordLogin(studentID, studentName, isReturning)
    
    ' 首次登录时添加水印和文件属性
    If isFirstLogin Then
        Call WatermarkManager.AddStudentWatermark(studentID, studentName)
    End If
    
    If isReturning Then
        ' 后续登录：只读模式
        MsgBox "欢迎回来 " & studentName & "！" & vbCrLf & vbCrLf & _
               "您已完成过作业，本次只能查看和阅读。" & vbCrLf & _
               "如需重新编辑，请联系老师重置。", vbInformation, "欢迎"
        
        ' 设置只读模式
        Call EnterReadOnlyMode(studentID, studentName)
    Else
        ' 首次登录：允许编辑
        MsgBox "欢迎 " & studentName & "！" & vbCrLf & vbCrLf & _
               "请完成作业，您可以编辑文档内容。" & vbCrLf & _
               "完成后请保存。", vbInformation, "欢迎"
        
        ' 移除保护，允许编辑
        On Error Resume Next
        ThisDocument.Unprotect
        If Err.Number <> 0 And Err.Number <> 5380 Then
            ThisDocument.Unprotect Password:="teacher2024"
        End If
        Err.Clear
        On Error GoTo 0
    End If
    
    g_IsLoggedIn = True
End Sub

' =====================================================
' 辅助功能
' =====================================================

' 获取已保存的学号（用于锁定验证）- 直接读取自定义属性
Private Function GetSavedStudentID() As String
    Dim props As Object
    Dim i As Long
    Dim prop As Object
    
    On Error Resume Next
    GetSavedStudentID = ""
    
    Set props = ThisDocument.CustomDocumentProperties
    If props Is Nothing Then Exit Function
    
    For i = 1 To props.Count
        Set prop = props(i)
        If prop.Name = "LockedStudentID" Then
            GetSavedStudentID = CStr(prop.Value)
            Exit For
        End If
    Next i
    
    On Error GoTo 0
End Function

' 获取已保存的密码（用于锁定验证）- 直接读取自定义属性
Private Function GetSavedStudentPwd() As String
    Dim props As Object
    Dim i As Long
    Dim prop As Object
    
    On Error Resume Next
    GetSavedStudentPwd = ""
    
    Set props = ThisDocument.CustomDocumentProperties
    If props Is Nothing Then Exit Function
    
    For i = 1 To props.Count
        Set prop = props(i)
        If prop.Name = "LockedStudentPwd" Then
            GetSavedStudentPwd = CStr(prop.Value)
            Exit For
        End If
    Next i
    
    On Error GoTo 0
End Function

' 保存学号和密码（首次登录时锁定）- 直接写入自定义属性
Private Sub SaveStudentCredentials(studentID As String, studentPwd As String)
    Dim props As Object
    Dim i As Long
    Dim prop As Object
    Dim doc As Document
    Dim bkName As String
    Dim rangeObj As Range
    
    Set doc = ThisDocument
    On Error Resume Next
    
    ' 1. 先删除旧的锁定书签（如果有）
    doc.Bookmarks("LockedStudent_ID").Delete
    doc.Bookmarks("LockedStudent_Pwd").Delete
    Err.Clear
    
    ' 2. 删除旧的同名属性（防止类型冲突）
    Set props = doc.CustomDocumentProperties
    If props.Count > 0 Then
        For i = props.Count To 1 Step -1
            If props(i).Name = "LockedStudentID" Or props(i).Name = "LockedStudentPwd" Then
                props(i).Delete
            End If
        Next i
    End If
    
    ' 3. 直接添加新的字符串属性
    props.Add Name:="LockedStudentID", _
              LinkToContent:=False, _
              Type:=msoPropertyTypeString, _
              Value:=studentID
    
    props.Add Name:="LockedStudentPwd", _
              LinkToContent:=False, _
              Type:=msoPropertyTypeString, _
              Value:=studentPwd
    
    ' 4. 创建书签（指向文档末尾的隐藏位置）
    Set rangeObj = doc.Content
    rangeObj.Collapse Direction:=wdCollapseEnd
    rangeObj.InsertBefore " "
    rangeObj.Font.Hidden = True
    bkName = "LockedStudent_ID"
    doc.Bookmarks.Add Name:=bkName, Range:=rangeObj
    
    On Error GoTo 0
End Sub

' 设置文档自定义属性

' 进入只读模式（禁止编辑）
Private Sub EnterReadOnlyMode(studentID As String, studentName As String)
    On Error Resume Next
    
    ' 解除之前的保护
    ThisDocument.Unprotect
    ThisDocument.Unprotect Password:="teacher2024"
    ThisDocument.Unprotect Password:="StudentReadOnly2024"
    ThisDocument.Unprotect Password:="NoSelect2024"
    
    ' 设置只读保护
    ThisDocument.Protect Password:="StudentReadOnly2024", _
                       Type:=wdAllowOnlyReading
    
    ' 添加水印（显示为只读）
    Call WatermarkManager.AddStudentWatermark(studentID & "(只读)", studentName)
    
    On Error GoTo 0
End Sub

' 禁止选择（通过保护文档为只读表单模式）
Public Sub DisableSelection()
    On Error Resume Next
    ' 先解除保护
    ThisDocument.Unprotect
    ' 保护文档，禁止选择和编辑
    ThisDocument.Protect Password:="NoSelect2024", _
                       Type:=wdAllowOnlyFormFields
    On Error GoTo 0
End Sub

' =====================================================
' 老师功能
' =====================================================

' 显示管理员功能菜单
Private Sub ShowAdminMenu()
    Dim choice As String
    
    choice = InputBox("请选择操作：" & vbCrLf & vbCrLf & _
                      "1. 清除学生锁定（重置文档）" & vbCrLf & _
                      "2. 查看所有学生记录" & vbCrLf & _
                      "3. 导出学生记录到文件" & vbCrLf & _
                      "4. 退出（不进行任何操作）" & vbCrLf & vbCrLf & _
                      "请输入序号（1-4）：", _
                      "管理员功能菜单", "4")
    
    Select Case choice
        Case "1"
            ' 清除学生锁定
            Call TeacherReset(True)
        Case "2"
            ' 查看学生记录
            Call ViewAllStudents
        Case "3"
            ' 导出记录
            Call ExportStudentRecords
        Case "4", ""
            ' 退出
    End Select
End Sub

' 老师重置功能 - 清除学生记录和锁定
' 可选参数 skipPassword: 管理员调用时跳过密码验证
Public Sub TeacherReset(Optional ByVal skipPassword As Boolean = False)
    Dim pwd As String
    Dim props As Object
    Dim i As Integer
    
    ' 如果不是管理员调用，则验证密码
    If Not skipPassword Then
        pwd = InputBox("请输入老师密码进行重置：", "老师重置")
        If pwd <> "Teacher2024" Then
            If pwd <> "" Then
                MsgBox "密码错误！", vbExclamation, "错误"
            End If
            Exit Sub
        End If
    End If
    
    ' 清除自定义属性中的学号和密码
    On Error Resume Next
    Set props = ThisDocument.CustomDocumentProperties
    If Not props Is Nothing Then
        For i = props.Count To 1 Step -1
            If props(i).Name = "LockedStudentID" Or props(i).Name = "LockedStudentPwd" Then
                props(i).Delete
            End If
        Next i
    End If
    ' 清除书签
    ThisDocument.Bookmarks("LockedStudent_ID").Delete
    ThisDocument.Bookmarks("LockedStudent_Pwd").Delete
    Err.Clear
    On Error GoTo 0
    
    ' 清除水印
    Call WatermarkManager.RemoveWatermark
    
    ' 清除学生记录
    Call StudentStorage.ClearAllStudentRecords
    
    MsgBox "学生记录和锁定已清除！", vbInformation, "重置成功"
End Sub

' =====================================================
' 清除所有文档属性（用于删除VBA前重置文档）
' 调用此函数后，文档将恢复到首次登录状态
' =====================================================
Public Sub ClearAllDocumentProperties()
    Dim props As Object
    Dim i As Integer
    Dim bk As Bookmark
    Dim doc As Document
    
    Set doc = ThisDocument
    
    On Error Resume Next
    
    ' 1. 清除所有自定义属性（包括学号、密码等）
    Set props = doc.CustomDocumentProperties
    If Not props Is Nothing Then
        For i = props.Count To 1 Step -1
            props(i).Delete
        Next i
    End If
    
    ' 2. 清除锁定相关书签
    doc.Bookmarks("LockedStudent_ID").Delete
    doc.Bookmarks("LockedStudent_Pwd").Delete
    Err.Clear
    
    ' 3. 清除水印
    Call WatermarkManager.RemoveWatermark
    
    ' 4. 清除学生记录
    Call StudentStorage.ClearAllStudentRecords
    
    ' 5. 解除文档保护
    doc.Unprotect
    doc.Unprotect Password:="teacher2024"
    doc.Unprotect Password:="StudentReadOnly2024"
    doc.Unprotect Password:="NoSelect2024"
    doc.Unprotect Password:="TempProtect2024"
    Err.Clear
    
    MsgBox "文档已恢复到首次登录状态！" & vbCrLf & vbCrLf & _
           "删除VBA代码后重新导入，将作为首次登录。", vbInformation, "重置成功"
    
    On Error GoTo 0
End Sub

' 查看所有学生记录
Public Sub ViewAllStudents()
    Dim students As Collection
    Dim student As Variant
    Dim info As String
    Dim bk As Bookmark
    Dim prefixLen As Long
    
    info = "已登录学生列表：" & vbCrLf & vbCrLf
    prefixLen = Len("Student_")
    
    On Error Resume Next
    
    For Each bk In ThisDocument.Bookmarks
        If Left(bk.Name, prefixLen) = "Student_" Then
            info = info & "- " & bk.Range.Text & vbCrLf
        End If
    Next
    
    If info = "已登录学生列表：" & vbCrLf & vbCrLf Then
        info = "暂无学生记录"
    End If
    
    MsgBox info, vbInformation, "学生记录"
    
    On Error GoTo 0
End Sub

' 导出学生记录
Public Sub ExportStudentRecords()
    Dim bk As Bookmark
    Dim prefixLen As Long
    Dim info As String
    Dim filePath As String
    Dim fileNum As Integer
    
    info = ""
    prefixLen = Len("Student_")
    
    On Error Resume Next
    
    For Each bk In ThisDocument.Bookmarks
        If Left(bk.Name, prefixLen) = "Student_" Then
            info = info & bk.Range.Text & vbCrLf
        End If
    Next
    
    If info = "" Then
        MsgBox "暂无学生记录！", vbInformation, "提示"
        Exit Sub
    End If
    
    ' 保存到文件
    filePath = ThisDocument.Path & "\学生登录记录.txt"
    
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, "学生作业登录记录"
    Print #fileNum, "=================="
    Print #fileNum, "导出时间：" & Format(Now, "yyyy-mm-dd hh:nn:ss")
    Print #fileNum, ""
    Print #fileNum, info
    Close #fileNum
    
    MsgBox "学生记录已导出到：" & vbCrLf & filePath, vbInformation, "导出成功"
    
    On Error GoTo 0
End Sub

