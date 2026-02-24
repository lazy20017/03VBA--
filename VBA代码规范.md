# VBA 代码撰写规则指南

## 1. 文件类型说明

| 文件扩展名 | 类型说明 | 用途 |
|-----------|---------|------|
| `.bas` | 标准模块 | 存放通用过程和函数 |
| `.cls` | 类模块 | 面向对象编程，定义类和对象 |
| `.frm` | 窗体模块 | 用户窗体和控件代码 |
| `.docm` / `.dotm` | 文档/模板模块 | 文档级事件代码 |

---

## 2. 事件命名规范

### 2.1 文档事件（ThisDocument）

```vb
' 文档打开时执行
Private Sub Document_Open()
    ' 文档打开时的处理逻辑
End Sub

' 另一种文档打开事件（兼容性更好）
Private Sub AutoOpen()
    ' 文档打开时的处理逻辑
End Sub
```

**注意**：
- `Document_Open` 和 `AutoOpen` 只需保留一个，建议使用 `Document_Open`
- 不要在事件中使用 `Application.OnTime` 延迟执行，可能导致类型错误
- 使用 `On Error Resume Next` 添加错误处理

### 2.2 工作簿/文档事件

| 事件 | 触发时机 |
|-----|---------|
| `Document_Open` | 文档打开时 |
| `AutoOpen` | 文档打开时（兼容性更好） |
| `Document_Close` | 文档关闭前 |
| `AutoClose` | 文档关闭前 |
| `NewDocument` | 新建文档时 |

---

## 3. 模块组织规范

### 3.1 标准模块 (.bas)

```vb
' ===============================
' 模块名称: LoginModule
' 功能描述: 处理用户登录验证
' 创建日期: 2024-01-01
' ===============================

Option Explicit

' --------------------------------
' 模块级变量
' --------------------------------
Dim m_LoginAttempts As Integer

' --------------------------------
' 公共过程
' --------------------------------

' 文档打开时自动运行
Sub Document_Open()
    On Error Resume Next
    ShowLoginForm
    If Err.Number <> 0 Then
        MsgBox "启动登录失败: " & Err.Description, vbCritical, "错误"
    End If
    On Error GoTo 0
End Sub

' 显示登录窗体
Sub ShowLoginForm()
    Dim StudentID As String
    Dim Name As String
    
    StudentID = InputBox("请输入学号:", "学生登录")
    If StudentID = "" Then Exit Sub
    
    Name = InputBox("请输入姓名:", "学生登录")
    If Name = "" Then Exit Sub
    
    If VerifyLogin(StudentID, Name) Then
        MsgBox "登录成功！", vbInformation, "欢迎"
    End If
End Sub

' --------------------------------
' 私有函数
' --------------------------------

Private Function VerifyLogin(StudentID As String, Name As String) As Boolean
    VerifyLogin = (Len(StudentID) > 0 And Len(Name) > 0)
End Function
```

### 3.2 类模块 (.cls)

```vb
' ===============================
' 类名称: DocumentGenerator
' 功能描述: 生成Word文档
' ===============================

Option Explicit

' 类属性
Private m_DocumentTitle As String
Private m_Author As String

' 构造函数
Private Sub Class_Initialize()
    m_Author = "System"
End Sub

' 析构函数
Private Sub Class_Terminate()
    ' 清理资源
End Sub

' 属性设置
Public Property Let Title(ByVal value As String)
    m_DocumentTitle = value
End Property

Public Property Get Title() As String
    Title = m_DocumentTitle
End Property

' 类方法
Public Sub GenerateDocument()
    ' 生成文档逻辑
End Sub
```

### 3.3 ThisDocument 模块

```vb
' ===============================
' ThisDocument - 文档级事件
' ===============================

Option Explicit

' 文档打开时执行
Private Sub Document_Open()
    On Error Resume Next
    ' 处理逻辑
    If Err.Number <> 0 Then
        Debug.Print "错误: " & Err.Description
    End If
    On Error GoTo 0
End Sub

' 文档关闭时执行
Private Sub Document_Close()
    On Error Resume Next
    ' 清理逻辑
    If Err.Number <> 0 Then
        Debug.Print "关闭错误: " & Err.Description
    End If
    On Error GoTo 0
End Sub
```

---

## 4. 命名规范

### 4.1 变量命名

| 类型 | 前缀 | 示例 |
|-----|------|-----|
| 字符串 | str | strName |
| 整数 | int | intCount |
| 长整数 | lng | lngTotal |
| 布尔值 | bln | blnIsValid |
| 对象 | obj | objDocument |
| 集合 | col | colItems |
| 数组 | arr | arrData |

### 4.2 过程/函数命名

- **公共过程**: `动词 + 名词` - `ShowLoginForm`, `SaveDocument`
- **私有过程**: `动词 + 名词` - `ProcessData`, `ValidateInput`
- **函数**: `Get/Is + 名词` - `GetUserName`, `IsValidLogin`

### 4.3 模块命名

- 使用有意义的英文名称
- 首字母大写
- 避免使用中文或特殊字符

---

## 5. 错误处理规范

### 5.1 基本错误处理

```vb
Sub Example()
    On Error Resume Next
    ' 可能出错的代码
    
    If Err.Number <> 0 Then
        MsgBox "错误: " & Err.Description, vbCritical
        Err.Clear
    End If
    On Error GoTo 0
End Sub
```

### 5.2 详细错误处理

```vb
Sub Example()
    On Error GoTo ErrorHandler
    
    ' 代码逻辑
    Exit Sub

ErrorHandler:
    MsgBox "错误编号: " & Err.Number & vbCrLf & _
           "错误描述: " & Err.Description & vbCrLf & _
           "错误来源: " & Err.Source, _
           vbCritical, "系统错误"
    
    ' 记录日志
    Debug.Print "错误发生位置: Example, 错误: " & Err.Description
    
    Resume Next
End Sub
```

---

## 6. 注释规范

### 6.1 模块头部注释

```vb
' ===============================
' 模块名称: XXX
' 功能描述: 
' 创建日期: YYYY-MM-DD
' 作者: 
' 版本: 1.0
' ===============================
```

### 6.2 过程/函数注释

```vb
' -------------------------------
' 过程名称: ProcessData
' 功能: 处理数据
' 参数: 
'   - strInput: 输入数据
' 返回值: 处理结果
' -------------------------------
```

### 6.3 代码内注释

```vb
' 声明变量
Dim strName As String

' 初始化
strName = ""

' 处理逻辑
If strName <> "" Then
    ' 执行操作
End If
```

---

## 7. 代码格式规范

### 7.1 缩进

- 使用 4 个空格缩进
- 保持代码层次清晰

### 7.2 行长度

- 单行代码尽量不超过 80 个字符
- 必要时使用续行符 `_`

### 7.3 空格使用

```vb
' 推荐
If x = 1 And y = 2 Then
    DoSomething
End If

' 不推荐
If x=1 And y=2 Then
    DoSomething
End If
```

---

## 8. 最佳实践

### 8.1 始终使用 Option Explicit

```vb
Option Explicit
Dim strName As String
```

### 8.2 避免使用宏自动执行延迟

```vb
' ❌ 不推荐 - 可能导致错误
Private Sub Document_Open()
    Application.OnTime Now + TimeValue("0:00:01"), "AutoRun"
End Sub

' ✅ 推荐 - 直接执行
Private Sub Document_Open()
    On Error Resume Next
    YourProcedure
    On Error GoTo 0
End Sub
```

### 8.3 释放对象

```vb
Sub Example()
    Dim objDoc As Object
    Set objDoc = Documents.Open("test.docm")
    
    ' 使用对象
    ' ...
    
    ' 清理
    If Not objDoc Is Nothing Then
        objDoc.Close SaveChanges:=False
        Set objDoc = Nothing
    End If
End Sub
```

### 8.4 使用 Const 定义常量

```vb
' 模块顶部定义
Const APP_NAME As String = "VBA工具"
Const MAX_RETRIES As Integer = 3
```

---

## 9. 常见错误及避免方法

### 9.1 TimeValue 格式错误

```vb
' ❌ 错误格式
TimeValue("00:00:00.5")

' ✅ 正确格式
TimeValue("0:00:00.5")
' 或使用 TimeSerial
TimeSerial(0, 0, 0.5)
```

### 9.2 事件重复定义

```vb
' ❌ 同时定义 Document_Open 和 AutoOpen
Private Sub Document_Open()
    ' 代码1
End Sub

Private Sub AutoOpen()
    ' 代码2 - 可能导致重复执行
End Sub

' ✅ 只保留一个
Private Sub Document_Open()
    ' 代码
End Sub
```

### 9.3 缺少对象检查

```vb
' ❌ 可能出错
objDoc.Close

' ✅ 建议先检查
If Not objDoc Is Nothing Then
    objDoc.Close
End If
```

---

## 10. 文件组织建议

```
项目文件夹/
├── 文档模板/
│   └── template.dotm
├── VBACODE/
│   ├── ThisDocument.bas      # 文档事件
│   ├── LoginModule.bas       # 登录模块
│   ├── DataProcess.bas       # 数据处理
│   └── DocumentGenerator.cls # 文档生成类
└── 资源/
    └── images/
```

---

## 附录: VBA 保留字

避免使用以下保留字作为变量名：

```
And, As, Boolean, ByRef, ByVal, Call, Case, CBool, CByte, CCur, 
CDate, CDbl, Chr, CInt, CLng, Const, CSng, CStr, Currency, CVar, 
Date, Debug, Declare, Dim, Do, Double, Each, Else, ElseIf, End, 
EndIf, Enum, Exit, False, For, Function, Get, Global, GoSub, GoTo, 
If, Imp, In, Integer, Is, Let, Lib, Like, Long, Loop, LSet, Me, 
Mod, New, Next, Not, Nothing, Object, On, Option, Optional, Or, 
ParamArray, Private, Property, Public, Put, ReDim, Resume, Return, 
Select Case, Set, Static, Step, Stop, String, Sub, Then, To, True, 
Type, TypeOf, UBound, Unlock, Until, Variant, With, WithEvents, 
Write#, Xor
```

---

**版本**: 1.0  
**创建日期**: 2026-02-24  
**最后更新**: 2026-02-24
