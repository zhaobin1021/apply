
Public Class Form1
    Private oWordApplication As Microsoft.Office.Interop.Word.Application
    Private oDocument As Microsoft.Office.Interop.Word.Document
    Private oRange As Microsoft.Office.Interop.Word.Range
    Private oSelection As Microsoft.Office.Interop.Word.Selection
    Private path
    Private newfilename
    'Public Sub New()
    '    '激活word接口
    '    oWordApplication = New Microsoft.Office.Interop.Word.Application
    '    oWordApplication.Visible = True
    'End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Dia As Object
        Dia = New OpenFileDialog
        Dia.title = "xxx"
        Dia.filter = "所有文件(*.*)|*.*|所有文件(*.*)|*.*"
        Dia.filterindex = 1
        Dia.restoredirectory = False
        If (Dia.showdialog() = DialogResult.OK) Then
            path = Dia.filename
        End If
        OpenFile(path)
        RichTextBox1.Text = oDocument.Range.Text
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim Dia As Object
        Dia = New SaveFileDialog
        Dia.title = "xxx"
        Dia.filter = "所有文件(*.*)|*.*|所有文件(*.*)|*.*"
        Dia.filterindex = 1
        Dia.filename = "文件未命名"
        Dia.restoredirectory = False
        If (Dia.showdialog() = DialogResult.OK) Then
            If Dia.CheckFileExists Then
                newfilename = Dia.filename + ".docx"
            Else
                newfilename = Dia.filename
            End If
            Me.NewDocWithModel(newfilename)
            End If
    End Sub
    '创建新文档
    Public Sub NewDocument()
        Dim missing = System.Reflection.Missing.Value
        Dim isVisible As Boolean = True
        oDocument = oWordApplication.Documents.Add(missing, missing, missing, missing)
        oDocument.Activate()
    End Sub
    '新建模板文档
    Public Sub NewDocWithModel(ByVal FileName As String)
        Dim missing = System.Reflection.Missing.Value
        Dim isVisible As Boolean = True
        Dim strName As String
        strName = FileName
        oWordApplication = New Microsoft.Office.Interop.Word.Application

        oDocument = oWordApplication.Documents.Add(strName, missing, missing, isVisible)
        oDocument.Content.InsertAfter(RichTextBox1.Text)

        oDocument.Activate()
    End Sub
    '打开word文件
    Public Sub OpenFile(ByVal FileName As String)
        Dim strName As String
        Dim isReadOnly As Boolean
        Dim isVisible As Boolean
        Dim missing = System.Reflection.Missing.Value
        oWordApplication = New Microsoft.Office.Interop.Word.Application
        strName = FileName
        isReadOnly = False
        isVisible = True
        oDocument = oWordApplication.Documents.Open(strName, missing, isReadOnly, missing, missing, missing, missing, missing, missing, missing, missing, isVisible, missing, missing, missing, missing)
        oDocument.Activate()
    End Sub
End Class
