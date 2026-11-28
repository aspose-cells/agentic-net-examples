// Title: Generate HTML from an Excel workbook and send it as a newsletter email using Aspose.Cells and System.Net.Mail in C#
// Common Searches: how to export an entire Excel workbook to HTML with embedded images using Aspose.Cells C# | send an HTML newsletter from a generated HTML string using System.Net.Mail in .NET | convert Excel to base64‑encoded HTML for email body in C# | C# code to load an Excel file, export to HTML and email via SSL‑enabled SMTP | Aspose.Cells HtmlSaveOptions settings for email‑ready HTML output
// Tags: Aspose.Cells export workbook to HTML with base64 images | C# create HTML email body from Excel using MemoryStream | System.Net.Mail send HTML newsletter via SSL SMTP | HtmlSaveOptions gridlines UTF‑8 encoding Aspose.Cells | SMTP client credentials configuration C# email sending

using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsEmailNewsletter
{
    // The example validates command‑line arguments, loads the specified Excel file with Aspose.Cells, exports the full workbook to an HTML string using HtmlSaveOptions (grid lines, base64‑encoded images, UTF‑8), builds a MailMessage with the HTML as the body and a plain‑text alternate view, and sends the newsletter via an SSL‑enabled SmtpClient configured with host, port and credentials.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Validate arguments
                if (args.Length < 8)
                {
                    Console.WriteLine("Usage:");
                    Console.WriteLine("  AsposeCellsEmailNewsletter.exe <ExcelPath> <RecipientEmail> <Subject> <SenderEmail> <SmtpHost> <SmtpPort> <SmtpUser> <SmtpPassword>");
                    return;
                }

                string excelPath = args[0];
                string recipientEmail = args[1];
                string emailSubject = args[2];
                string senderEmail = args[3];
                string smtpHost = args[4];
                int smtpPort = int.Parse(args[5]);
                string smtpUser = args[6];
                string smtpPassword = args[7];

                // Verify that the Excel file exists
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"Error: The Excel file \"{excelPath}\" does not exist.");
                    return;
                }

                Workbook workbook;
                try
                {
                    // Load the workbook
                    workbook = new Workbook(excelPath);
                }
                catch (Exception loadEx)
                {
                    Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                    return;
                }

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportGridLines = true,               // Show grid lines in the HTML
                    ExportImagesAsBase64 = true,          // Embed images directly in the HTML
                    ExportActiveWorksheetOnly = false,    // Export the whole workbook
                    Encoding = Encoding.UTF8               // Ensure UTF‑8 encoding
                };

                // Save the workbook to a memory stream as HTML
                string htmlBody;
                try
                {
                    using (MemoryStream htmlStream = new MemoryStream())
                    {
                        workbook.Save(htmlStream, htmlOptions);
                        htmlBody = Encoding.UTF8.GetString(htmlStream.ToArray());
                    }
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to convert workbook to HTML: {saveEx.Message}");
                    return;
                }

                // Create the email message
                using (MailMessage mail = new MailMessage())
                {
                    mail.From = new MailAddress(senderEmail);
                    mail.To.Add(new MailAddress(recipientEmail));
                    mail.Subject = emailSubject;
                    mail.Body = htmlBody;
                    mail.IsBodyHtml = true;

                    // Optional: add a plain‑text alternate view for clients that do not support HTML
                    AlternateView plainView = AlternateView.CreateAlternateViewFromString(
                        "Your email client does not support HTML content.",
                        Encoding.UTF8,
                        "text/plain");
                    mail.AlternateViews.Add(plainView);

                    // Configure and send via SMTP
                    using (SmtpClient smtpClient = new SmtpClient(smtpHost, smtpPort))
                    {
                        smtpClient.EnableSsl = true;
                        smtpClient.Credentials = new NetworkCredential(smtpUser, smtpPassword);
                        smtpClient.Send(mail);
                    }
                }

                Console.WriteLine("Newsletter email sent successfully.");
            }
            catch (FormatException fe)
            {
                Console.WriteLine($"Argument format error: {fe.Message}");
            }
            catch (SmtpException se)
            {
                Console.WriteLine($"SMTP error: {se.StatusCode} - {se.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.GetType().Name} - {ex.Message}");
            }
        }
    }
}
