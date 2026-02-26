' -*- coding: utf-8 -*-
' =====================================================
' 学生信息存储模块
' 功能：记录学生登录信息，判断是否为Returning学生
' 存储位置：文档自定义XML属性或隐藏书签
' =====================================================

Option Explicit

' Word 常量声明
Const wdCollapseEnd As Long = 0

' 存储键名常量
Private Const STUDENT_INFO_BOOKMARK As String = "StudentLoginInfo"
Private Const STUDENT_DATA_PREFIX As String = "Student_"

' 学生信息类
Private Type StudentInfo
    StudentID As String
    StudentName As String
    LoginCount As Long
    FirstLoginDate As String
    LastLoginDate As String
    IsCompleted As Boolean
End Type

' 记录学生登录
Public Sub RecordLogin(studentID As String, studentName As String, isReturning As Boolean)
    Dim student As StudentInfo
    Dim bookmark As Bookmark
    Dim found As Boolean
    
    ' 查找是否已有该学生的记录
    student = GetStudentInfo(studentID)
    found = (student.StudentID <> "")
    
    If found Then
        ' 更新已有记录
        student.LoginCount = student.LoginCount + 1
        student.LastLoginDate = Format(Now, "yyyy-mm-dd hh:nn:ss")
        If Not isReturning Then
            student.IsCompleted = True
        End If
    Else
        ' 创建新记录
        student.StudentID = studentID
        student.StudentName = studentName
        student.LoginCount = 1
        student.FirstLoginDate = Format(Now, "yyyy-mm-dd hh:nn:ss")
        student.LastLoginDate = Format(Now, "yyyy-mm-dd hh:nn:ss")
        student.IsCompleted = isReturning  ' 首次登录时isReturning为False
    End If
    
    ' 保存学生信息
    Call SaveStudentInfo(student)
End Sub

' 检查是否为Returning学生
Public Function IsReturningStudent(studentID As String, studentName As String) As Boolean
    Dim student As StudentInfo
    
    student = GetStudentInfo(studentID)
    
    ' 如果找到记录且已完成过作业，则为Returning学生
    If student.StudentID <> "" Then
        ' 检查姓名是否匹配
        If Trim(student.StudentName) = Trim(studentName) Then
            IsReturningStudent = student.IsCompleted Or student.LoginCount > 1
        Else
            IsReturningStudent = False
        End If
    Else
        IsReturningStudent = False
    End If
End Function

' 获取学生信息
Private Function GetStudentInfo(studentID As String) As StudentInfo
    Dim student As StudentInfo
    Dim dataStr As String
    Dim parts() As String
    
    On Error Resume Next
    
    ' 尝试从书签获取数据
    Dim bk As Bookmark
    Set bk = ThisDocument.Bookmarks(STUDENT_DATA_PREFIX & studentID)
    
    If Not bk Is Nothing Then
        dataStr = bk.Range.Text
        
        ' 解析数据：学号|姓名|登录次数|首次登录时间|最后登录时间|是否完成
        parts = Split(dataStr, "|")
        
        If UBound(parts) >= 5 Then
            student.StudentID = parts(0)
            student.StudentName = parts(1)
            student.LoginCount = CLng(parts(2))
            student.FirstLoginDate = parts(3)
            student.LastLoginDate = parts(4)
            student.IsCompleted = (parts(5) = "1")
        End If
    End If
    
    GetStudentInfo = student
    
    On Error GoTo 0
End Function

' 保存学生信息
Private Sub SaveStudentInfo(student As StudentInfo)
    Dim dataStr As String
    Dim bookmark As Bookmark
    Dim bk As Bookmark
    
    On Error Resume Next
    
    ' 构建数据字符串：学号|姓名|登录次数|首次登录时间|最后登录时间|是否完成
    dataStr = student.StudentID & "|" & _
              student.StudentName & "|" & _
              CStr(student.LoginCount) & "|" & _
              student.FirstLoginDate & "|" & _
              student.LastLoginDate & "|" & _
              IIf(student.IsCompleted, "1", "0")
    
    ' 使用自定义属性存储学生信息（不显示在文档中）
    Call SetStudentProperty(student.StudentID, "Data", dataStr)
    
    On Error GoTo 0
End Sub

' 设置学生自定义属性
Private Sub SetStudentProperty(studentID As String, propSuffix As String, propValue As String)
    Dim props As Object
    Dim propName As String
    Dim i As Integer
    Dim found As Boolean
    Dim strValue As String
    
    On Error Resume Next
    
    propName = "Student_" & studentID & "_" & propSuffix
    Set props = ThisDocument.CustomDocumentProperties
    
    If props Is Nothing Then
        Exit Sub
    End If
    
    ' 安全处理：确保传入的是字符串类型
    If IsEmpty(propValue) Or IsNull(propValue) Then
        strValue = ""
    ElseIf VarType(propValue) = vbBoolean Then
        strValue = CStr(propValue)
    Else
        strValue = CStr(propValue)
    End If
    
    ' 检查是否已存在该属性
    found = False
    For i = 1 To props.Count
        If props(i).Name = propName Then
            props(i).Value = strValue
            found = True
            Exit For
        End If
    Next i
    
    ' 如果不存在，则添加新属性
    If Not found Then
        ThisDocument.CustomDocumentProperties.Add Name:=propName, _
            LinkToContent:=False, _
            Type:=2, _
            Value:=strValue
    End If
    
    On Error GoTo 0
End Sub

' 获取学生自定义属性
Private Function GetStudentProperty(studentID As String, propSuffix As String) As String
    Dim props As Object
    Dim propName As String
    Dim i As Integer
    Dim propValue As Variant
    
    GetStudentProperty = ""
    
    On Error Resume Next
    
    propName = "Student_" & studentID & "_" & propSuffix
    Set props = ThisDocument.CustomDocumentProperties
    
    If props Is Nothing Then
        Exit Function
    End If
    
    For i = 1 To props.Count
        If props(i).Name = propName Then
            propValue = props(i).Value
            
            ' 安全处理：确保返回的是字符串类型，而不是Boolean等
            If IsEmpty(propValue) Or IsNull(propValue) Then
                GetStudentProperty = ""
            ElseIf VarType(propValue) = vbBoolean Then
                GetStudentProperty = CStr(propValue)
            ElseIf VarType(propValue) = vbString Then
                GetStudentProperty = CStr(propValue)
            Else
                GetStudentProperty = CStr(propValue)
            End If
            
            Exit For
        End If
    Next i
    
    On Error GoTo 0
End Function

' 删除学生自定义属性
Private Sub DeleteStudentProperty(studentID As String, propSuffix As String)
    Dim props As Object
    Dim propName As String
    Dim i As Integer
    
    On Error Resume Next
    
    propName = "Student_" & studentID & "_" & propSuffix
    Set props = ThisDocument.CustomDocumentProperties
    
    If props Is Nothing Then
        Exit Sub
    End If
    
    For i = props.Count To 1 Step -1
        If props(i).Name = propName Then
            props(i).Delete
            Exit For
        End If
    Next i
    
    On Error GoTo 0
End Sub

' 获取所有已登录学生列表
Public Function GetAllStudents() As Collection
    Dim students As New Collection
    Dim bk As Bookmark
    Dim prefixLen As Long
    
    prefixLen = Len(STUDENT_DATA_PREFIX)
    
    On Error Resume Next
    
    For Each bk In ThisDocument.Bookmarks
        If Left(bk.Name, prefixLen) = STUDENT_DATA_PREFIX Then
            students.Add bk.Name
        End If
    Next
    
    Set GetAllStudents = students
    
    On Error GoTo 0
End Function

' 清除所有学生记录（老师功能）
Public Sub ClearAllStudentRecords()
    Dim bk As Bookmark
    Dim toDelete As Collection
    Dim bkName As Variant
    
    Set toDelete = New Collection
    
    On Error Resume Next
    
    ' 收集要删除的书签
    For Each bk In ThisDocument.Bookmarks
        If Left(bk.Name, Len(STUDENT_DATA_PREFIX)) = STUDENT_DATA_PREFIX Then
            toDelete.Add bk.Name
        End If
    Next
    
    ' 删除书签
    For Each bkName In toDelete
        ThisDocument.Bookmarks(CStr(bkName)).Delete
    Next
    
    MsgBox "已清除所有学生记录！", vbInformation, "提示"
    
    On Error GoTo 0
End Sub

