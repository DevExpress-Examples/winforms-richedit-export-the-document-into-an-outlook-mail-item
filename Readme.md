<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/128610249/13.1.4%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/E4438)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
[![](https://img.shields.io/badge/💬_Leave_Feedback-feecdd?style=flat-square)](#does-this-example-address-your-development-requirementsobjectives)
<!-- default badges end -->

# Rich Text Editor for WinForms - How to Export the RichEditControl Document into an Outlook Mail Item

The following example describes how to send the mail merge result as an e-mail.

## Implementation Details

In some scenarios, you might wish just to transfer the RichEditControl document into Outlook. In this case, use [Outlook Interop API](https://learn.microsoft.com/en-us/office/client-developer/outlook/pia/welcome-to-the-outlook-primary-interop-assembly-reference?redirectedfrom=MSDN) to prepare a mail item based on the RichEditControl content. Process images via a custom [IUriProvider](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Office.Services.IUriProvider) interface implementation. Convert native RichEdit images into Outlook mail item attachments.

Refer to the following web articles to learn how to deal with the Outlook-related part of this solution:

* [How to embed image in HTML body in c# into Outlook mail](https://learn.microsoft.com/en-us/archive/msdn-technet-forums/6c063b27-7e8a-4963-ad5f-ce7e5ffb2c64)
* [How to embed images in email](https://stackoverflow.com/questions/4312687/how-to-embed-images-in-email)

## Files to Review

* [Form1.cs](./CS/Form1.cs) (VB: [Form1.vb](./VB/Form1.vb))
* [Program.cs](./CS/Program.cs) (VB: [Program.vb](./VB/Program.vb))

## More Examples

* [Rich Text Editor for WinForms - Build a Mail Application with the RichEditControl](https://github.com/DevExpress-Examples/build-a-mail-application-with-the-richeditcontrol)

## Documentation

* [How to: Send the Mail-Merge Document as an E-Mail](https://docs.devexpress.com/WindowsForms/120456/controls-and-libraries/rich-text-editor/examples/import-and-export/how-to-send-the-mail-merge-document-as-an-e-mail)
<!-- feedback -->
## Does this example address your development requirements/objectives?

[<img src="https://www.devexpress.com/support/examples/i/yes-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=winforms-richedit-export-the-document-into-an-outlook-mail-item&~~~was_helpful=yes) [<img src="https://www.devexpress.com/support/examples/i/no-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=winforms-richedit-export-the-document-into-an-outlook-mail-item&~~~was_helpful=no)

(you will be redirected to DevExpress.com to submit your response)
<!-- feedback end -->
