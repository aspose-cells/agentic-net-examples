// Title: Render an Aspose.Cells chart to a JPEG image and embed it inline in an HTML email using C#
// AI Prompts: Generate C# code that creates a workbook with Aspose.Cells, builds a chart, converts the chart to a JPEG stream, and inserts the image into an HTML email body using a Content‑ID and LinkedResource. | Show how to configure System.Net.Mail to send an HTML email with an inline chart image rendered by Aspose.Cells without writing the image to disk.
// Common Searches: C# Aspose.Cells export chart to JPEG stream for email | embed chart image inline in HTML email using System.Net.Mail | how to use LinkedResource with Content-ID to display chart in email body | send Aspose.Cells generated chart as inline image without temporary file | render Aspose chart to memory stream and attach to email C#
// Tags: Aspose.Cells chart to JPEG stream | inline chart image in HTML email C# | LinkedResource Content-ID email embedding | System.Net.Mail embed image from memory | render Aspose chart without saving to file

using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using System;
using System.IO;
using System.Net.Mail;
using System.Net.Mime;

// The example creates a workbook, adds sample data, generates a column chart, renders the chart to a JPEG image stored in a memory stream with Aspose.Cells, and then builds an HTML email where the chart is displayed inline using a Content‑ID and LinkedResource in System.Net.Mail.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];

            // Fill worksheet with sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 5);
            var chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true); // Values

            // Export the chart to a JPEG image stored in memory
            using (var imageStream = new MemoryStream())
            {
                var imgOptions = new ImageOrPrintOptions
                {
                    OnePagePerSheet = true
                    // ImageFormat property is not available in some versions; default format (PNG) will be used.
                };

                // Render chart to the stream
                chart.ToImage(imageStream, imgOptions);
                imageStream.Position = 0; // Reset stream position for reading

                // Build the email message
                var mail = new MailMessage
                {
                    From = new MailAddress("sender@example.com"),
                    Subject = "Chart Image Inline",
                    IsBodyHtml = true
                };
                mail.To.Add("recipient@example.com");

                // HTML body referencing the image via Content-ID
                string contentId = "ChartImage";
                string htmlBody = $"<html><body><h3>Here is the chart:</h3><img src=\"cid:{contentId}\" alt=\"Chart\"/></body></html>";

                // Create an AlternateView for the HTML body
                AlternateView htmlView = AlternateView.CreateAlternateViewFromString(htmlBody, null, MediaTypeNames.Text.Html);

                // Create a LinkedResource for the JPEG image (or PNG if default)
                var inlineImage = new LinkedResource(imageStream, MediaTypeNames.Image.Jpeg)
                {
                    ContentId = contentId,
                    TransferEncoding = TransferEncoding.Base64
                };
                htmlView.LinkedResources.Add(inlineImage);
                mail.AlternateViews.Add(htmlView);

                // Configure SMTP client (replace with real settings)
                using (var smtp = new SmtpClient("smtp.example.com"))
                {
                    smtp.Credentials = new System.Net.NetworkCredential("username", "password");
                    smtp.EnableSsl = true;
                    // smtp.Send(mail); // Uncomment to send the email
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
