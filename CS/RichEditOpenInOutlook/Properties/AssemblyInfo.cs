// Developer Express Code Central Example:
// How to export the RichEditControl document into an Outlook mail item
// 
// We have the http://www.devexpress.com/scid=E2216 example that illustrates how to
// create a self-contained email client application based on our RichEditControl
// (http://documentation.devexpress.com/#WindowsForms/CustomDocument6975). Note
// that only the System.Net.Mail functionality is used in this example and the
// message is sent directly from the RichEditControl. However, in some scenarios
// (e.g., see http://www.devexpress.com/scid=Q423631), you might wish just to
// transfer the RichEditControl document into Outlook. In this case, use Outlook
// Interop API (http://msdn.microsoft.com/en-us/library/office/bb652780.aspx) to
// prepare a mail item based on the RichEditControl content. Process images via a
// custom IUriProvider Interface
// (http://documentation.devexpress.com/#CoreLibraries/clsDevExpressXtraRichEditServicesIUriProvidertopic)
// implementor. Convert native RichEdit images into Outlook mail item attachments.
// Refer to the following web articles to learn how to deal with the
// Outlook-related part of this solution:
// how to embed image in html body in c#
// into outlook mail
// (http://social.msdn.microsoft.com/Forums/en-US/vsto/thread/6c063b27-7e8a-4963-ad5f-ce7e5ffb2c64/)
// Attach
// stream data with Outlook mail client
// (http://social.msdn.microsoft.com/Forums/pl/outlookdev/thread/17efe46b-18fe-450f-9f6e-d8bb116161d8)
// How
// to embed images in email
// (http://stackoverflow.com/questions/4312687/how-to-embed-images-in-email)
// 
// You can find sample updates and versions for different programming languages here:
// http://www.devexpress.com/example=E4438

using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

// General Information about an assembly is controlled through the following 
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
[assembly: AssemblyTitle("RichEditOpenInOutlook")]
[assembly: AssemblyDescription("")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("Microsoft")]
[assembly: AssemblyProduct("RichEditOpenInOutlook")]
[assembly: AssemblyCopyright("Copyright © Microsoft 2010")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

// Setting ComVisible to false makes the types in this assembly not visible 
// to COM components.  If you need to access a type in this assembly from 
// COM, set the ComVisible attribute to true on that type.
[assembly: ComVisible(false)]

// The following GUID is for the ID of the typelib if this project is exposed to COM
[assembly: Guid("68427cc0-f04e-465f-803a-e7c4ecc1438a")]

// Version information for an assembly consists of the following four values:
//
//      Major Version
//      Minor Version 
//      Build Number
//      Revision
//
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]
