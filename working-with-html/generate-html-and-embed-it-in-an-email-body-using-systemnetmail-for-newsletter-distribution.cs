// Title: Generate a self‑contained HTML newsletter from an Aspose.Cells workbook and send it via System.Net.Mail in C#
// AI Prompts: Write C# code that creates an Aspose.Cells workbook, exports it to a self‑contained HTML string (embedding images as Base64), and sets that string as the HTML body of a MailMessage for SMTP delivery. | Extend the sample to add an alternate plain‑text view to the MailMessage while preserving the HTML view generated from the workbook.
// Common Searches: c# convert aspose.cells workbook to html string with base64 images for email | how to send aspose.cells generated html as email body using System.Net.Mail | embed excel data as html newsletter in smtp client c# | asp.net core send html email with embedded spreadsheet content aspose cells | self contained html email from excel workbook c# example
// Tags: Aspose.Cells HTML export with base64 images | C# send HTML email using System.Net.Mail | self-contained HTML email body | convert workbook to HTML string | SMTP client newsletter with embedded spreadsheet data

using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Text;
using Aspose.Cells;

// The program builds a simple workbook with product data, converts it to a self‑contained HTML string using Aspose.Cells HtmlSaveOptions (including Base64‑encoded images), and sends that HTML as the body of an email via System.Net.Mail with SSL/TLS authentication.
class NewsletterSender
{
    static void Main()
    {
        // 1. Create a new workbook and populate it with sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Sample data
        cells["A1"].PutValue("Product");
        cells["B1"].PutValue("Price");
        cells["A2"].PutValue("Apple");
        cells["B2"].PutValue(1.20);
        cells["A3"].PutValue("Banana");
        cells["B3"].PutValue(0.80);
        cells["A4"].PutValue("Cherry");
        cells["B4"].PutValue(2.50);

        // 2. Convert the workbook to HTML and store it in a string
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            // Ensure the HTML is self‑contained (styles embedded)
            ExportImagesAsBase64 = true,
            ExportActiveWorksheetOnly = true
        };

        string htmlContent;
        using (MemoryStream htmlStream = new MemoryStream())
        {
            workbook.Save(htmlStream, htmlOptions);
            htmlStream.Position = 0;
            using (StreamReader reader = new StreamReader(htmlStream, Encoding.UTF8))
            {
                htmlContent = reader.ReadToEnd();
            }
        }

        // 3. Prepare the email message
        MailMessage mail = new MailMessage
        {
            From = new MailAddress("sender@example.com"),
            Subject = "Monthly Newsletter",
            Body = htmlContent,
            IsBodyHtml = true // Important: tells the client that the body is HTML
        };

        // Add recipients (example)
        mail.To.Add("recipient1@example.com");
        mail.To.Add("recipient2@example.com");

        // 4. Configure the SMTP client
        SmtpClient smtp = new SmtpClient
        {
            Host = "smtp.example.com",      // SMTP server address
            Port = 587,                     // Common port for TLS
            EnableSsl = true,               // Use SSL/TLS
            Credentials = new NetworkCredential("smtp_user", "smtp_password")
        };

        try
        {
            // 5. Send the email
            smtp.Send(mail);
            Console.WriteLine("Newsletter sent successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error sending newsletter: " + ex.Message);
        }
        finally
        {
            // Clean up resources
            mail.Dispose();
            smtp.Dispose();
        }
    }
}
