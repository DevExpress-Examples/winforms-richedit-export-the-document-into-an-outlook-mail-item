using DevExpress.Office.Services;
using DevExpress.Office.Utils;
using DevExpress.Utils;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.Export;
using System;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace RichEditOpenInOutlook
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            richEdit.LoadDocument("Hello.docx");
        }

        private void btnSend_Click(object sender, EventArgs e)
        {
            if ((edtTo.Text.Trim() == "") || (edtSubject.Text.Trim() == ""))
            {
                MessageBox.Show("Fill in required fields");
                return;
            }
            try
            {
                Type outlookType = Type.GetTypeFromProgID("Outlook.Application");
                if (outlookType == null)
                {
                    MessageBox.Show("Microsoft Outlook is not installed.");
                    return;
                }

                object application = Activator.CreateInstance(outlookType);
                object mailItem = InvokeMethod(application, "CreateItem", 0);

                SetProperty(mailItem, "To", edtTo.Text);
                SetProperty(mailItem, "Subject", edtSubject.Text);

                RichEditMailMessageExporter exporter = new RichEditMailMessageExporter(richEdit, mailItem);
                exporter.Export();

                InvokeMethod(mailItem, "Display", false);
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
            }
        }

        static object InvokeMethod(object target, string methodName, params object[] args)
        {
            return target.GetType().InvokeMember(methodName, BindingFlags.InvokeMethod, null, target, args);
        }

        static object GetProperty(object target, string propertyName)
        {
            return target.GetType().InvokeMember(propertyName, BindingFlags.GetProperty, null, target, null);
        }

        static void SetProperty(object target, string propertyName, object value)
        {
            target.GetType().InvokeMember(propertyName, BindingFlags.SetProperty, null, target, new[] { value });
        }

        public class RichEditMailMessageExporter : IUriProvider
        {
            readonly RichEditControl control;
            readonly object mailItem;
            int imageId;
            string tempFiles = Path.Combine(Directory.GetCurrentDirectory(), "TempFiles");

            public RichEditMailMessageExporter(RichEditControl control, object mailItem)
            {
                Guard.ArgumentNotNull(control, "control");
                Guard.ArgumentNotNull(mailItem, "mailItem");

                this.control = control;
                this.mailItem = mailItem;
            }

            public virtual void Export()
            {
                if (!Directory.Exists(tempFiles))
                    Directory.CreateDirectory(tempFiles);

                control.BeforeExport += OnBeforeExport;
                string htmlBody = control.Document.GetHtmlText(control.Document.Range, this);
                control.BeforeExport -= OnBeforeExport;

                SetProperty(mailItem, "BodyFormat", 2);
                SetProperty(mailItem, "HTMLBody", htmlBody);
            }

            private void OnBeforeExport(object sender, BeforeExportEventArgs e)
            {
                HtmlDocumentExporterOptions options = e.Options as HtmlDocumentExporterOptions;
                if (options != null)
                {
                    options.Encoding = Encoding.UTF8;
                }
            }

            #region IUriProvider Members
            public string CreateCssUri(string rootUri, string styleText, string relativeUri)
            {
                return String.Empty;
            }

            public string CreateImageUri(string rootUri, OfficeImage image, string relativeUri)
            {
                string imageName = String.Format("image{0}.png", imageId);
                imageId++;

                string imagePath = Path.Combine(tempFiles, imageName);

                image.NativeImage.Save(imagePath, ImageFormat.Png);

                object attachments = GetProperty(mailItem, "Attachments");
                InvokeMethod(attachments, "Add", imagePath, 1, 0, Type.Missing);

                return "cid:" + imageName;
            }
            #endregion
        }
    }
}