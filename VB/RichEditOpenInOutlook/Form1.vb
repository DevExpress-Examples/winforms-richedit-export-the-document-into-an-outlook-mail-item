' Developer Express Code Central Example:
' How to export the RichEditControl document into an Outlook mail item
' 
' We have the http://www.devexpress.com/scid=E2216 example that illustrates how to
' create a self-contained email client application based on our RichEditControl
' (http://documentation.devexpress.com/#WindowsForms/CustomDocument6975). Note
' that only the System.Net.Mail functionality is used in this example and the
' message is sent directly from the RichEditControl. However, in some scenarios
' (e.g., see http://www.devexpress.com/scid=Q423631), you might wish just to
' transfer the RichEditControl document into Outlook. In this case, use Outlook
' Interop API (http://msdn.microsoft.com/en-us/library/office/bb652780.aspx) to
' prepare a mail item based on the RichEditControl content. Process images via a
' custom IUriProvider Interface
' (http://documentation.devexpress.com/#CoreLibraries/clsDevExpressXtraRichEditServicesIUriProvidertopic)
' implementor. Convert native RichEdit images into Outlook mail item attachments.
' Refer to the following web articles to learn how to deal with the
' Outlook-related part of this solution:
' how to embed image in html body in c#
' into outlook mail
' (http://social.msdn.microsoft.com/Forums/en-US/vsto/thread/6c063b27-7e8a-4963-ad5f-ce7e5ffb2c64/)
' Attach
' stream data with Outlook mail client
' (http://social.msdn.microsoft.com/Forums/pl/outlookdev/thread/17efe46b-18fe-450f-9f6e-d8bb116161d8)
' How
' to embed images in email
' (http://stackoverflow.com/questions/4312687/how-to-embed-images-in-email)
' 
' You can find sample updates and versions for different programming languages here:
' http://www.devexpress.com/example=E4438

Imports System
Imports System.IO
Imports System.Text
Imports System.Windows.Forms
Imports System.Drawing.Imaging
Imports DevExpress.Utils
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.Export
Imports DevExpress.XtraRichEdit.Utils
Imports DevExpress.Office.Utils
Imports DevExpress.Office.Services
Imports Outlook = Microsoft.Office.Interop.Outlook
Imports DevExpress.XtraRichEdit.Services

Namespace RichEditOpenInOutlook
	Partial Public Class Form1
		Inherits Form

		Public Sub New()
			InitializeComponent()

			richEdit.LoadDocument("Hello.docx")
		End Sub

		Private Sub btnSend_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSend.Click
			If (edtTo.Text.Trim() = "") OrElse (edtSubject.Text.Trim() = "") Then
				MessageBox.Show("Fill in required fields")
				Return
			End If
			Try
				Dim application As New Outlook.Application()
				Dim mailItem As Outlook.MailItem = DirectCast(application.CreateItem(Outlook.OlItemType.olMailItem), Outlook.MailItem)

				mailItem.To = edtTo.Text
				mailItem.Subject = edtSubject.Text

				Dim exporter As New RichEditMailMessageExporter(richEdit, mailItem)
				exporter.Export()

				mailItem.Display(False)
			Catch exc As Exception
				MessageBox.Show(exc.Message)
			End Try
		End Sub

		Public Class RichEditMailMessageExporter
			Implements IUriProvider

			Private ReadOnly control As RichEditControl
			Private ReadOnly mailItem As Outlook.MailItem
			Private imageId As Integer
			Private tempFiles As String = Path.Combine(Directory.GetCurrentDirectory(), "TempFiles")

			Public Sub New(ByVal control As RichEditControl, ByVal mailItem As Outlook.MailItem)
				Guard.ArgumentNotNull(control, "control")
				Guard.ArgumentNotNull(mailItem, "mailItem")

				Me.control = control
				Me.mailItem = mailItem
			End Sub

			Public Overridable Sub Export()
				If Not Directory.Exists(tempFiles) Then
					Directory.CreateDirectory(tempFiles)
				End If

				AddHandler control.BeforeExport, AddressOf OnBeforeExport
				Dim htmlBody As String = control.Document.GetHtmlText(control.Document.Range, Me)
				RemoveHandler control.BeforeExport, AddressOf OnBeforeExport

				mailItem.BodyFormat = Outlook.OlBodyFormat.olFormatHTML
				mailItem.HTMLBody = htmlBody
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

				mailItem.Attachments.Add(imagePath, Outlook.OlAttachmentType.olByValue, 0, Type.Missing)

				Return "cid:" & imageName
			End Function
			#End Region
		End Class
	End Class
End Namespace