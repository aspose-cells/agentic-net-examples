// Title: Embed Aspose.Cells Chart as Inline JPEG in HTML Email (C#)
// Description: Creates a workbook, builds a column chart, converts the chart to a JPEG, encodes it as a Base64 data URI, and sends an HTML email with the chart displayed inline using SmtpClient.
// Keywords: Aspose.Cells | chart to JPEG | inline image email | Base64 data URI | C# SmtpClient | embed chart in email | Excel chart image | HTML email embedding
// Common Searches: Aspose.Cells embed chart in email C# | C# send chart as inline image email | convert Excel chart to JPEG base64 | inline JPEG in HTML email C# | Aspose.Cells chart to Base64 for email
// Developer Intent: Generate a JPEG snapshot of an Aspose.Cells chart and embed it directly in the HTML body of an email without using attachments.
// Use Cases: Automated daily reports that show chart previews inside the email body. | Alert notifications that present key metrics as embedded charts. | Batch emails containing multiple chart images inline for comprehensive dashboards.
// AI Prompts: Generate C# code that converts an Aspose.Cells chart to a PNG and embeds it inline in an email using a Base64 data URI. | Explain how to embed several Aspose.Cells chart images as separate inline images within one HTML email. | Show how to replace the temporary file with an in‑memory stream for chart image conversion and email embedding.

using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, builds a column chart, converts the chart to a JPEG, encodes it as a Base64 data URI, and sends an HTML email with the chart displayed inline using SmtpClient.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["A3"].PutValue("Banana");
            worksheet.Cells["A4"].PutValue("Cherry");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(30);
            worksheet.Cells["B3"].PutValue(45);
            worksheet.Cells["B4"].PutValue(25);

            // Add a column chart and set its data source
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Save chart as JPEG to a temporary file
            string tempImagePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".jpg");
            try
            {
                chart.ToImage(tempImagePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error converting chart to image: {ex.Message}");
                return;
            }

            // Ensure the image file was created
            if (!File.Exists(tempImagePath))
            {
                Console.WriteLine("Chart image file was not created.");
                return;
            }

            // Read image bytes and encode as Base64 for inline HTML
            byte[] imageBytes = File.ReadAllBytes(tempImagePath);
            string base64Image = Convert.ToBase64String(imageBytes);
            string imgSrc = $"data:image/jpeg;base64,{base64Image}";

            // Build the HTML body with the embedded image
            string htmlBody = $"<html><body><h2>Embedded Chart</h2><img src=\"{imgSrc}\" alt=\"Chart\"/></body></html>";

            // Configure the email message
            using (MailMessage mail = new MailMessage())
            {
                mail.From = new MailAddress("sender@example.com");
                mail.To.Add("recipient@example.com");
                mail.Subject = "Chart Image Embedded in Email";
                mail.Body = htmlBody;
                mail.IsBodyHtml = true;

                // Set up the SMTP client (replace with actual server details)
                using (SmtpClient smtp = new SmtpClient("smtp.example.com", 587))
                {
                    smtp.Credentials = new NetworkCredential("username", "password");
                    smtp.EnableSsl = true;

                    try
                    {
                        smtp.Send(mail);
                        Console.WriteLine("Email sent successfully.");
                    }
                    catch (SmtpException ex)
                    {
                        Console.WriteLine($"SMTP error: {ex.Message}");
                    }
                }
            }

            // Clean up temporary image file
            try
            {
                File.Delete(tempImagePath);
            }
            catch
            {
                // Ignore any errors during cleanup
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
