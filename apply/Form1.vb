﻿Imports Microsoft.Office.Interop
Imports mshtml

Public Class Form1
    Private oWordApplication As Word.Application
    Private oDocument As Word.Document
    Private oRange As Word.Range
    Private oSelection As Word.Selection
    Private html As mshtml.HTMLDocument
    Private path
    Private newfilename
    Private ofd
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
        newfilename = "C:\Users\zhaobin\Desktop\文档.docx"
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
            'tgtFile = "C:\Users\zhaobin\Desktop\文档图片.html"
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
        Dim arr As Array
        arr = sender.split(".")
        toHtml = arr.GetValue(0) + ".html"
        Return toHtml
    End Function

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        '  Dim strName As String
        '  Dim isReadOnly As Boolean
        '  Dim isVisible As Boolean
        ' Dim document
        ' Dim docu
        ' Dim missing = System.Reflection.Missing.Value
        ' WebBrowser.Url =
        ' oWordApplication = New Word.Application
        'strName = "file:///C:\Users\zhaobin\Desktop\A20J792000-C8-1-1a.html"

        'Shell("D:\360\360se6\Application\360se.exe")
        WebBrowser1.Url = New Uri(String.Format("file:///C:\Users\zhaobin\Desktop\A20J792000-C8-1-1a.html", Application.StartupPath))

        'Button1.Enabled = False
        'html = New HTMLDocument
        ' docu = WebBrowser1.Document.Body.InnerHtml
        'docu = WebBrowser1.Document.Body.GetElementsByTagName("span")

        'document = html.open(strName)
        'document = html.(strName)
        'html.getElementsByTagName("//span")


        ''strName = tgtFile
        'isReadOnly = False
        'isVisible = True
        '' oDocument = oWordApplication.Documents.Open(strName, missing, isReadOnly, missing, missing, missing, missing, missing, missing, missing, missing, isVisible, missing, missing, missing, missing)
        'oDocument = oWordApplication.Documents.Open(strName)

        ''Dim reader As IO.StreamReader = New IO.StreamReader(, System.Text.Encoding.GetEncoding("GB2312"))
        'Dim respHTML As String = oDocument.ReadToEnd()

        ' RichTextBox1.Text = docu
        'oDocument.Activate()
        'oDocument.Close()
        'oWordApplication.Quit()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

    End Sub

    Private Sub WebBrowser1_DragOver(sender As Object, e As DragEventArgs) Handles WebBrowser1.DragOver
        RichTextBox1.Text = WebBrowser1.Document.Body.InnerHtml
    End Sub
End Class
