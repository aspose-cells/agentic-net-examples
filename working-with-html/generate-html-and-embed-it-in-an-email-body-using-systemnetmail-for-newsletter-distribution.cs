// Title: Convert Aspose.Cells Worksheet to HTML5 with Base64 Images and Send via System.Net.Mail (C# Newsletter)
// Description: Creates an Excel workbook with Aspose.Cells, exports the used range as HTML5 (images embedded as Base64), converts the output to a UTF‑8 string, and sends it as the HTML body of an email using System.Net.Mail's SmtpClient. Includes basic error handling and SMTP configuration.
// Keywords: Aspose.Cells HTML export | HTML5 email body C# | Base64 images in email | System.Net.Mail newsletter | Convert Excel range to HTML | C# send HTML email | Aspose.Cells ToHtml | SMTP client C#
// Common Searches: how to export Aspose.Cells worksheet to HTML string | send HTML email with embedded images using System.Net.Mail | C# convert Excel range to HTML5 with Base64 | Aspose.Cells newsletter email example | C# email body from Excel data
// Developer Intent: Generate an HTML5 representation of an Excel range (including Base64‑encoded images) and deliver it as the body of an email via System.Net.Mail.
// Use Cases: Monthly newsletter generated from Excel data | Automated sales or inventory report emailed as styled HTML | Product catalog distribution where Excel tables become web‑ready email content | Sending chart‑rich Excel dashboards in email without attachments
// AI Prompts: Show C# code to export a specific Aspose.Cells range to an HTML string with Base64 images and send it via System.Net.Mail. | Explain how to add multiple attachments while keeping the HTML body intact. | Provide best‑practice error‑handling and retry logic for SMTP delivery after converting Excel to HTML.

using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Text;
using Aspose.Cells;

// Creates an Excel workbook with Aspose.Cells, exports the used range as HTML5 (images embedded as Base64), converts the output to a UTF‑8 string, and sends it as the HTML body of an email using System.Net.Mail's SmtpClient. Includes basic error handling and SMTP configuration.
class NewsletterEmail
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Welcome to our Newsletter");
            sheet.Cells["A2"].PutValue("Here is some data:");
            for (int i = 1; i <= 5; i++)
            {
                sheet.Cells[i + 2, 0].PutValue($"Item {i}");
                sheet.Cells[i + 2, 1].PutValue(i * 10);
            }

            // Initialize HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportImagesAsBase64 = true,          // Embed images as Base64
                HtmlVersion = HtmlVersion.Html5      // Use HTML5 standard
            };

            // Convert the used range of the worksheet to HTML bytes
            Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
            byte[] htmlBytes = usedRange.ToHtml(saveOptions);
            string htmlBody = Encoding.UTF8.GetString(htmlBytes); // Convert bytes to string

            // Prepare the email message
            using (MailMessage mail = new MailMessage())
            {
                mail.From = new MailAddress("sender@example.com");
                mail.To.Add("recipient@example.com");
                mail.Subject = "Monthly Newsletter";
                mail.Body = htmlBody;
                mail.IsBodyHtml = true; // Indicate that the body contains HTML

                // Configure the SMTP client (replace with actual server details)
                using (SmtpClient smtp = new SmtpClient("smtp.example.com", 587))
                {
                    smtp.Credentials = new NetworkCredential("username", "password");
                    smtp.EnableSsl = true;

                    // Send the email
                    smtp.Send(mail);
                }
            }
        }
        catch (Exception ex)
        {
            // Log or handle exceptions as needed
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
