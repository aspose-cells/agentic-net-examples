// Title: Send an Aspose.Cells worksheet as mobile‑friendly HTML email using System.Net.Mail (C#)
// Description: Creates a workbook, fills a sheet with data, converts the used range to a Base64‑embedded, mobile‑compatible HTML string via HtmlSaveOptions, and sends it as the body of an SMTP email with System.Net.Mail.
// Keywords: Aspose.Cells HTML export C# | Convert worksheet to HTML string | ExportImagesAsBase64 | IsMobileCompatible HTML email | System.Net.Mail HTML body | C# email newsletter from Excel | SMTP send HTML email Aspose | Excel data in email without attachment
// Common Searches: how to convert Aspose.Cells range to HTML string C# | send Excel data as HTML email using System.Net.Mail | Aspose.Cells HtmlSaveOptions for email newsletters | embed Base64 images in HTML email from workbook | C# code to email spreadsheet as HTML
// Developer Intent: Generate an HTML representation of a workbook range and deliver it directly in the body of an email via SMTP.
// Use Cases: Weekly sales dashboard emailed as a responsive HTML table to subscribers. | Automated product‑catalog newsletter that embeds chart images without attachments. | Transactional order confirmation that displays order details from an Excel sheet in the email body.
// AI Prompts: Write C# code that loads an Aspose.Cells workbook, converts a specific range to a Base64‑embedded, mobile‑compatible HTML string, and sends it with System.Net.Mail using UTF‑8 encoding. | Explain the HtmlSaveOptions settings needed for email‑ready HTML and how to attach the generated HTML to a MailMessage.

using System;
using System.Net;
using System.Net.Mail;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsEmailDemo
{
    // Creates a workbook, fills a sheet with data, converts the used range to a Base64‑embedded, mobile‑compatible HTML string via HtmlSaveOptions, and sends it as the body of an SMTP email with System.Net.Mail.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and populate it with sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                for (int row = 0; row < 10; row++)
                {
                    for (int col = 0; col < 5; col++)
                    {
                        sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                    }
                }

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportImagesAsBase64 = true,   // embed images if any
                    IsMobileCompatible = true     // mobile friendly
                };

                // Convert the used range of the worksheet to HTML (as byte[])
                Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
                byte[] htmlBytes = usedRange.ToHtml(htmlOptions);

                // Convert the byte array to a UTF‑8 string
                string htmlBody = Encoding.UTF8.GetString(htmlBytes);

                // Prepare the email message
                MailMessage message = new MailMessage
                {
                    From = new MailAddress("sender@example.com"),
                    Subject = "Workbook as HTML Email",
                    IsBodyHtml = true,
                    Body = htmlBody
                };
                message.To.Add("recipient@example.com");

                // Configure the SMTP client (replace with real host/credentials)
                using (SmtpClient smtp = new SmtpClient("smtp.example.com", 587))
                {
                    smtp.EnableSsl = true;
                    smtp.Credentials = new NetworkCredential("smtp_user", "smtp_password");

                    // Send the email
                    smtp.Send(message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
