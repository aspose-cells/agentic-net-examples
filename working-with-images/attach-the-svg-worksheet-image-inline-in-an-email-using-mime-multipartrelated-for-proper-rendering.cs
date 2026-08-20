// Title: Aspose.Cells for .NET: Render a Worksheet to SVG and Embed It Inline in an Email (MIME multipart/related)
// Description: This C# example shows how to create a workbook with Aspose.Cells, render the first worksheet page to an SVG image using SheetRender and SvgImageOptions, load the SVG into a MemoryStream, and build a MIME multipart/related email. The SVG is added as a LinkedResource with Content‑ID "WorksheetSvg" and referenced in the HTML body via an <img> tag, ready for delivery through SmtpClient.
// Keywords: Aspose.Cells | C# | SVG rendering | SheetRender | SvgImageOptions | inline email image | MIME multipart related | LinkedResource | Content-ID | System.Net.Mail | SMTP | US developers | European developers | India developers
// Common Searches: render excel worksheet to svg c# | embed svg in email using asp.net | aspnet send inline svg image email | mime multipart related email c# example | linkedresource svg image asp.net mailmessage | aspose.cells svg email tutorial
// Developer Intent: Create an SVG snapshot of an Excel worksheet and embed it directly in the email body as an inline image.
// Use Cases: Automated sales dashboards sent as SVG previews in daily report emails. | Embedding live worksheet visuals in marketing newsletters without separate attachments. | Sending monitoring alerts with instant SVG charts for quick data interpretation.
// AI Prompts: Generate C# code that uses Aspose.Cells to convert a worksheet to SVG and embed the SVG inline in an email using AlternateView and LinkedResource. | Explain step‑by‑step how to configure a MIME multipart/related email with an inline SVG image, including Content‑ID handling and client compatibility. | Provide troubleshooting tips when an inline SVG does not render in Outlook, Gmail, or Apple Mail after being sent.

using System;
using System.IO;
using System.Net.Mail;
using System.Net.Mime;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// This C# example shows how to create a workbook with Aspose.Cells, render the first worksheet page to an SVG image using SheetRender and SvgImageOptions, load the SVG into a MemoryStream, and build a MIME multipart/related email. The SVG is added as a LinkedResource with Content‑ID "WorksheetSvg" and referenced in the HTML body via an <img> tag, ready for delivery through SmtpClient.
class SvgEmailExample
{
    static void Main()
    {
        try
        {
            // 1. Create a workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["B3"].PutValue(95);

            // 2. Configure SVG rendering options
            SvgImageOptions svgOptions = new SvgImageOptions
            {
                // FitToViewPort = true // optional: fit to viewport
            };

            // 3. Render the first worksheet page to an SVG file
            string svgPath = Path.Combine(Path.GetTempPath(), "worksheet.svg");
            SheetRender renderer = new SheetRender(sheet, svgOptions);
            renderer.ToImage(0, svgPath); // renders page 0 to the specified file

            // 4. Load the generated SVG into a memory stream (ensure the file exists)
            if (!File.Exists(svgPath))
                throw new FileNotFoundException("SVG file was not created.", svgPath);

            byte[] svgBytes = File.ReadAllBytes(svgPath);
            using (MemoryStream svgStream = new MemoryStream(svgBytes))
            {
                // 5. Prepare the email with an inline SVG image
                MailMessage message = new MailMessage
                {
                    From = new MailAddress("sender@example.com"),
                    Subject = "Worksheet as Inline SVG"
                };
                message.To.Add("recipient@example.com");

                // HTML body referencing the SVG via Content-ID
                string htmlBody = @"<html><body>
                                    <h2>Worksheet Snapshot</h2>
                                    <img src=""cid:WorksheetSvg"" alt=""Worksheet SVG"" />
                                    </body></html>";

                // Create an AlternateView for HTML content
                AlternateView htmlView = AlternateView.CreateAlternateViewFromString(
                    htmlBody, null, MediaTypeNames.Text.Html);

                // Create a LinkedResource for the SVG image
                LinkedResource svgResource = new LinkedResource(svgStream, "image/svg+xml")
                {
                    ContentId = "WorksheetSvg",
                    TransferEncoding = TransferEncoding.Base64,
                    ContentType = { MediaType = "image/svg+xml" }
                };

                // Attach the SVG as a linked resource
                htmlView.LinkedResources.Add(svgResource);
                message.AlternateViews.Add(htmlView);

                // 6. Send the email (SMTP settings should be configured appropriately)
                using (SmtpClient smtp = new SmtpClient("smtp.example.com", 587))
                {
                    smtp.Credentials = new System.Net.NetworkCredential("username", "password");
                    smtp.EnableSsl = true;

                    // Uncomment the line below to actually send the email
                    // smtp.Send(message);
                }
            }

            // Clean up temporary SVG file
            if (File.Exists(svgPath))
            {
                File.Delete(svgPath);
            }

            Console.WriteLine("Email prepared with inline SVG.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
