Imports DevExpress.Office.Services
Imports DevExpress.Office.Utils
Imports DevExpress.Utils
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.Export
Imports System
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Reflection
Imports System.Text
Imports System.Windows.Forms

Namespace RichEditOpenInOutlook

    Public Partial Class Form1
        Inherits Form

        Public Sub New()
            InitializeComponent()
            richEdit.LoadDocument("Hello.docx")
        End Sub

        Private Sub btnSend_Click(ByVal sender As Object, ByVal e As EventArgs)
            If(Equals(edtTo.Text.Trim(), "")) OrElse (Equals(edtSubject.Text.Trim(), "")) Then
                MessageBox.Show("Fill in required fields")
                Return
            End If

            Try
                Dim outlookType As Type = Type.GetTypeFromProgID("Outlook.Application")
                If outlookType Is Nothing Then
                    MessageBox.Show("Microsoft Outlook is not installed.")
                    Return
                End If

                Dim application As Object = Activator.CreateInstance(outlookType)
                Dim mailItem As Object = InvokeMethod(application, "CreateItem", 0)
                SetProperty(mailItem, "To", edtTo.Text)
                SetProperty(mailItem, "Subject", edtSubject.Text)
                Dim exporter As RichEditMailMessageExporter = New RichEditMailMessageExporter(richEdit, mailItem)
                exporter.Export()
                InvokeMethod(mailItem, "Display", False)
            Catch exc As Exception
                MessageBox.Show(exc.Message)
            End Try
        End Sub

        Private Shared Function InvokeMethod(ByVal target As Object, ByVal methodName As String, ParamArray args As Object()) As Object
            Return target.GetType().InvokeMember(methodName, BindingFlags.InvokeMethod, Nothing, target, args)
        End Function

        Private Shared Function GetProperty(ByVal target As Object, ByVal propertyName As String) As Object
            Return target.GetType().InvokeMember(propertyName, BindingFlags.GetProperty, Nothing, target, Nothing)
        End Function

        Private Shared Sub SetProperty(ByVal target As Object, ByVal propertyName As String, ByVal value As Object)
            target.GetType().InvokeMember(propertyName, BindingFlags.SetProperty, Nothing, target, {value})
        End Sub

        Public Class RichEditMailMessageExporter
            Implements IUriProvider

            Private ReadOnly control As RichEditControl

            Private ReadOnly mailItem As Object

            Private imageId As Integer

            Private tempFiles As String = Path.Combine(Directory.GetCurrentDirectory(), "TempFiles")

            Public Sub New(ByVal control As RichEditControl, ByVal mailItem As Object)
                Guard.ArgumentNotNull(control, "control")
                Guard.ArgumentNotNull(mailItem, "mailItem")
                Me.control = control
                Me.mailItem = mailItem
            End Sub

            Public Overridable Sub Export()
                If Not Directory.Exists(tempFiles) Then Directory.CreateDirectory(tempFiles)
                AddHandler control.BeforeExport, AddressOf OnBeforeExport
                Dim htmlBody As String = control.Document.GetHtmlText(control.Document.Range, Me)
                RemoveHandler control.BeforeExport, AddressOf OnBeforeExport
                SetProperty(mailItem, "BodyFormat", 2)
                SetProperty(mailItem, "HTMLBody", htmlBody)
            End Sub

            Private Sub OnBeforeExport(ByVal sender As Object, ByVal e As BeforeExportEventArgs)
                Dim options As HtmlDocumentExporterOptions = TryCast(e.Options, HtmlDocumentExporterOptions)
                If options IsNot Nothing Then
                    options.Encoding = Encoding.UTF8
                End If
            End Sub

#Region "IUriProvider Members"
            Public Function CreateCssUri(ByVal rootUri As String, ByVal styleText As String, ByVal relativeUri As String) As String Implements IUriProvider.CreateCssUri
                Return String.Empty
            End Function

            Public Function CreateImageUri(ByVal rootUri As String, ByVal image As OfficeImage, ByVal relativeUri As String) As String Implements IUriProvider.CreateImageUri
                Dim imageName As String = String.Format("image{0}.png", imageId)
                imageId += 1
                Dim imagePath As String = Path.Combine(tempFiles, imageName)
                image.NativeImage.Save(imagePath, ImageFormat.Png)
                Dim attachments As Object = GetProperty(mailItem, "Attachments")
                InvokeMethod(attachments, "Add", imagePath, 1, 0, Type.Missing)
                Return "cid:" & imageName
            End Function
#End Region
        End Class
    End Class
End Namespace
