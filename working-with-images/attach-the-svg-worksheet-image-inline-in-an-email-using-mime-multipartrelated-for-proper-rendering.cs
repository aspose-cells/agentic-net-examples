// Title: Convert an Excel worksheet to SVG and embed it as an inline image in a multipart/related email using C# and Aspose.Cells
// AI Prompts: Write C# code that loads an .xlsx workbook with Aspose.Cells, renders the first worksheet to an SVG MemoryStream, creates a LinkedResource with a Content-ID, builds an HTML AlternateView that references the CID, and sends the email via SmtpClient. | Show how to configure ImageOrPrintOptions for SVG output, use SheetRender to generate the SVG stream, and attach it as an inline image in a multipart/related email using .NET's MailMessage and LinkedResource classes.
// Common Searches: how to embed a worksheet rendered as SVG in an HTML email using C# | C# Aspose.Cells convert Excel sheet to SVG and send as inline image | using LinkedResource to embed SVG in multipart/related email in .NET | render Excel worksheet to SVG stream for email body C# | send SVG image inline in email with SmtpClient and Aspose.Cells
// Tags: Aspose.Cells SVG worksheet rendering | C# LinkedResource inline SVG email | multipart/related email with embedded SVG | SheetRender to MemoryStream C# | SMTP email sending with embedded worksheet image

using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using System.Drawing.Imaging;

// The example loads an Excel workbook, renders the first worksheet to an SVG image using Aspose.Cells, creates a LinkedResource with a CID for the SVG stream, builds an HTML body that references the CID, adds the resource to a multipart/related AlternateView, and sends the email through SmtpClient.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the Excel workbook
            var workbook = new Workbook(inputPath);

            // Configure rendering options for SVG output
            var imgOptions = new ImageOrPrintOptions
            {
                SaveFormat = SaveFormat.Svg,   // Correct property for SVG format
                OnePagePerSheet = true
            };

            // Render the first worksheet to SVG and store it in a memory stream
            var sheetRender = new SheetRender(workbook.Worksheets[0], imgOptions);
            using (var svgStream = new MemoryStream())
            {
                // Render the first page of the sheet (index 0) to SVG
                sheetRender.ToImage(0, svgStream);
                svgStream.Position = 0; // Reset stream position for reading

                // Create a linked resource for the SVG image
                var svgResource = new LinkedResource(svgStream, new ContentType("image/svg+xml"))
                {
                    ContentId = "worksheetSvg",
                    TransferEncoding = TransferEncoding.Base64
                };

                // Build the HTML body that references the SVG via CID
                string htmlBody = @"<html><body>
                                    <h2>Worksheet as SVG</h2>
                                    <img src=""cid:worksheetSvg"" alt=""Worksheet SVG"" />
                                    </body></html>";

                // Create an alternate view with multipart/related content type
                var htmlView = AlternateView.CreateAlternateViewFromString(htmlBody, null, MediaTypeNames.Text.Html);
                htmlView.LinkedResources.Add(svgResource);

                // Prepare the email message
                var mail = new MailMessage
                {
                    From = new MailAddress("sender@example.com"),
                    Subject = "Worksheet SVG Inline Image",
                    IsBodyHtml = true
                };
                mail.To.Add("recipient@example.com");
                mail.AlternateViews.Add(htmlView);

                // Send the email (configure SMTP as needed)
                using (var smtp = new SmtpClient("smtp.example.com", 587))
                {
                    smtp.Credentials = new NetworkCredential("smtp_user", "smtp_password");
                    smtp.EnableSsl = true;
                    smtp.Send(mail);
                }
            }
        }
        catch (Exception ex)
        {
            // Log or display the exception details for troubleshooting
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
