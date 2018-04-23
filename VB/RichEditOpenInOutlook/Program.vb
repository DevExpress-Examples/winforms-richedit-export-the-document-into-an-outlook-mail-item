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
Imports System.Collections.Generic
Imports System.Windows.Forms

Namespace RichEditOpenInOutlook
	Friend NotInheritable Class Program

		Private Sub New()
		End Sub

		''' <summary>
		''' The main entry point for the application.
		''' </summary>
		<STAThread>
		Shared Sub Main()
			Application.EnableVisualStyles()
			Application.SetCompatibleTextRenderingDefault(False)
			Application.Run(New Form1())
		End Sub
	End Class
End Namespace