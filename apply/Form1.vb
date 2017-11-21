Imports Microsoft.Office.Interop
Public Class Form1
    Private oWordApplication As Word.Application
    Private oDocument As Word.Document
    Private oRange As Word.Range
    Private oSelection As Word.Selection
    Private path
    Private newfilename
    Private ofd As New OpenFileDialog
    Private tgtFile As Object
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
        '模板可以自己选择
        'Dim Dia As Object
        'Dia = New SaveFileDialog
        'Dia.title = "xxx"
        'Dia.filter = "所有文件(*.*)|*.*|所有文件(*.*)|*.*"
        'Dia.filterindex = 1
        'Dia.filename = "文件未命名"
        'Dia.restoredirectory = False
        'If (Dia.showdialog() = DialogResult.OK) Then
        '    If Dia.CheckFileExists Then
        '        newfilename = Dia.filename + ".docx"
        '    Else
        '        newfilename = Dia.filename
        '    End If
        '    Me.NewDocWithModel(newfilename)
        'End If

        '默认模板
        newfilename = "C:\Users\zhaobin\Desktop\文件未命名.docx"
        Me.NewDocWithModel(newfilename)
    End Sub
    '创建新模板文档
    Public Sub NewDocument()
        Dim missing = System.Reflection.Missing.Value
        Dim isVisible As Boolean = True
        oDocument = oWordApplication.Documents.Add(missing, missing, missing, missing)

        oDocument.Activate()
        oDocument.Close()
    End Sub
    '新建模板文档,通过已有文档模板建立新文档
    Public Sub NewDocWithModel(ByVal FileName As String)
        Dim missing = System.Reflection.Missing.Value
        Dim isVisible As Boolean = True
        Dim strName As String
        strName = FileName
        oWordApplication = New Word.Application
        oDocument = oWordApplication.Documents.Add(strName, missing, missing, isVisible)
        'oDocument.Shapes.AddPicture("C:\Users\zhaobin\Desktop\1.png")
        oDocument.Content.InsertAfter(RichTextBox1.Text)
        oDocument.ActiveWindow.WindowState = 1
        oWordApplication.Application.Visible = True
        oDocument.Activate()
    End Sub
    '打开word文件
    Public Sub OpenFile(ByVal FileName As String)
        Dim strName As String
        Dim isReadOnly As Boolean
        Dim isVisible As Boolean
        Dim missing = System.Reflection.Missing.Value
        oWordApplication = New Word.Application
        strName = FileName
        isReadOnly = False
        isVisible = True
        oDocument = oWordApplication.Documents.Open(strName, missing, isReadOnly, missing, missing, missing, missing, missing, missing, missing, missing, isVisible, missing, missing, missing, missing)
        oDocument.Activate()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If ofd.ShowDialog() = DialogResult.OK Then
            tgtFile = ofd.FileName
            oWordApplication = New Word.Application
            oDocument = oWordApplication.Documents.Open(tgtFile)
            tgtFile = "C:\Users\zhaobin\Desktop\文档图片.html"
            tgtFile = WordToTtml(tgtFile)
            'oDocument.SaveAs2(tgtFile, FileFormat:="wdFormatHTML")格式错误
            oDocument.SaveAs2(tgtFile, Word.WdSaveFormat.wdFormatHTML)
            'oDocument.SaveAs(tgtFile, Word.WdSaveFormat.wdFormatHTML)
            'saveas2比saveas多一个参数CompatibilityMode，兼容模式
            oDocument.Close()
            oWordApplication.Quit()
        End If
    End Sub
    Function WordToTtml(sender As Object)
        Dim toHtml
        For i = 0 To sender.ToString.Length

        Next
        Return toHtml
    End Function
End Class
