' ========================================
' DocumentGenerator Class - Generate document with Student ID, Name and Timestamp
' ========================================

Private p_StudentID As String
Private p_Name As String
Private p_Time As Date

' Property: StudentID
Public Property Get StudentID() As String
    StudentID = p_StudentID
End Property

Public Property Let StudentID(Value As String)
    p_StudentID = Value
End Property

' Property: Name
Public Property Get Name() As String
    Name = p_Name
End Property

Public Property Let Name(Value As String)
    p_Name = Value
End Property

' Property: Time
Public Property Get Time() As Date
    Time = p_Time
End Property

Public Property Let Time(Value As Date)
    p_Time = Value
End Property

' Constructor
Private Sub Class_Initialize()
    p_Time = Now()
End Sub

' Generate document
Public Sub GenerateStudentDocument(OutputPath As String)
    ' Create new document
    Dim doc As Document
    Set doc = Documents.Add
    
    ' Set document title
    doc.Content.InsertBefore "Student Information Document" & vbCrLf
    doc.Paragraphs(1).Range.Font.Size = 16
    doc.Paragraphs(1).Range.Font.Bold = True
    doc.Paragraphs(1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    
    ' Add student information
    Dim table As Table
    Set table = doc.Tables.Add(Range:=doc.Paragraphs(2).Range, _
                              NumRows:=3, NumColumns:=2)
    
    ' Set table style
    table.Borders.Enable = True
    table.Style = "Grid Table"
    
    ' Fill table content
    ' Row 1: Student ID
    table.Cell(1, 1).Range.Text = "Student ID"
    table.Cell(1, 2).Range.Text = p_StudentID
    
    ' Row 2: Name
    table.Cell(2, 1).Range.Text = "Name"
    table.Cell(2, 2).Range.Text = p_Name
    
    ' Row 3: Time
    table.Cell(3, 1).Range.Text = "Record Time"
    table.Cell(3, 2).Range.Text = Format(p_Time, "yyyy-mm-dd HH:mm:ss")
    
    ' Add signature area
    doc.Paragraphs.Add
    doc.Content.InsertAfter vbCrLf & "============= Signature Area =============" & vbCrLf & vbCrLf
    
    ' Student signature line
    doc.Content.InsertAfter "Student Signature:___________    " & "Date:___________" & vbCrLf & vbCrLf
    
    ' Teacher signature line
    doc.Content.InsertAfter "Teacher Signature:___________    " & "Date:___________" & vbCrLf
    
    ' Save document
    doc.SaveAs2 OutputPath
    
    MsgBox "Document generated: " & OutputPath, vbInformation, "Success"
End Sub

' Export to Excel (optional feature)
Public Sub ExportToExcel(StudentID As String, Name As String)
    Dim xlApp As Object
    Dim xlWorkbook As Object
    Dim xlSheet As Object
    
    ' Create Excel application
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True
    
    ' Create workbook
    Set xlWorkbook = xlApp.Workbooks.Add
    Set xlSheet = xlWorkbook.Sheets(1)
    
    ' Write data
    xlSheet.Cells(1, 1).Value = "Student ID"
    xlSheet.Cells(1, 2).Value = "Name"
    xlSheet.Cells(1, 3).Value = "Time"
    
    xlSheet.Cells(2, 1).Value = StudentID
    xlSheet.Cells(2, 2).Value = Name
    xlSheet.Cells(2, 3).Value = Now()
    
    ' Auto-fit columns
    xlSheet.Columns("A:C").AutoFit
    
    MsgBox "Data exported to Excel", vbInformation, "Success"
End Sub
