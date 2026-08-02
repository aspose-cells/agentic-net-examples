// Title: Embed an Aspose.Cells Worksheet as Inline SVG in a C# Email (MIME multipart/related)
// Description: This example creates a Workbook, fills it with data, renders the first worksheet to an in‑memory SVG using Aspose.Cells SvgImageOptions, and builds a MailMessage with an HTML body that references the SVG via a Content‑ID. The SVG is added as a LinkedResource (image/svg+xml) and the complete multipart/related message is saved as an .eml file through an SmtpClient configured for a pickup directory.
// Keywords: Aspose.Cells | C# | SVG rendering | inline SVG email | multipart/related | LinkedResource | MailMessage eml | SmtpClient pickup directory | worksheet to SVG | email attachment without file
// Common Searches: how to embed SVG in .NET email | Aspose.Cells render worksheet to SVG | C# send email with inline image | save MailMessage as .eml file | multipart related email example C# | inline worksheet snapshot email
// Developer Intent: Create an email that displays a worksheet as an inline SVG image, without requiring external attachments, and store the message as an .eml file.
// Use Cases: Send daily sales reports as a visual SVG snapshot embedded directly in the email body. | Archive compliance‑required notifications with embedded worksheet graphics for easy review. | Automate alerts that include a rendered worksheet diagram, enabling recipients to see data instantly without opening attachments.
// AI Prompts: Generate C# code that converts an Aspose.Cells worksheet to an SVG stream and embeds it in a MailMessage using multipart/related. | Explain how to configure SmtpClient to write a MailMessage containing an inline SVG to a .eml file via a pickup directory. | Provide error‑handling best practices for rendering SVG with Aspose.Cells and attaching it as a LinkedResource in .NET emails.

using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// This example creates a Workbook, fills it with data, renders the first worksheet to an in‑memory SVG using Aspose.Cells SvgImageOptions, and builds a MailMessage with an HTML body that references the SVG via a Content‑ID. The SVG is added as a LinkedResource (image/svg+xml) and the complete multipart/related message is saved as an .eml file through an SmtpClient configured for a pickup directory.
class SvgWorksheetEmailExample
{
    static void Main()
    {
        try
        {
            // 1. Create a workbook and populate it with sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["B3"].PutValue(85);

            // 2. Configure SVG rendering options (no ImageFormat property needed)
            SvgImageOptions svgOptions = new SvgImageOptions
            {
                FitToViewPort = true // Fit the SVG to the viewport
            };

            // 3. Render the worksheet to an in‑memory SVG stream
            using (MemoryStream svgStream = new MemoryStream())
            {
                SheetRender renderer = new SheetRender(sheet, svgOptions);
                renderer.ToImage(0, svgStream); // Render first (and only) page
                svgStream.Position = 0; // Reset stream for reading

                // 4. Build the email message with a multipart/related body
                MailMessage message = new MailMessage
                {
                    From = new MailAddress("sender@example.com"),
                    Subject = "Worksheet as Inline SVG"
                };
                message.To.Add("recipient@example.com");

                // HTML body referencing the embedded SVG via Content‑ID
                string htmlBody = @"
                    <html>
                    <body>
                        <h2>Worksheet Snapshot</h2>
                        <img src=""cid:worksheetSvg"" alt=""Worksheet SVG""/>
                    </body>
                    </html>";

                // Create an alternate view for HTML content
                AlternateView htmlView = AlternateView.CreateAlternateViewFromString(
                    htmlBody, null, MediaTypeNames.Text.Html);

                // Attach the SVG as a linked resource
                LinkedResource svgResource = new LinkedResource(svgStream, "image/svg+xml")
                {
                    ContentId = "worksheetSvg",
                    TransferEncoding = TransferEncoding.Base64
                };
                htmlView.LinkedResources.Add(svgResource);
                message.AlternateViews.Add(htmlView);

                // 5. Write the MIME message to a file using a pickup directory
                string emlPath = "WorksheetEmail.eml";
                string pickupDir = Path.GetDirectoryName(Path.GetFullPath(emlPath)) ?? Directory.GetCurrentDirectory();
                Directory.CreateDirectory(pickupDir); // Ensure directory exists

                try
                {
                    using (FileStream emlFile = new FileStream(emlPath, FileMode.Create))
                    {
                        // The .NET MailMessage does not provide direct saving,
                        // so we use SmtpClient with a specified pickup directory.
                        SmtpClient pickupClient = new SmtpClient
                        {
                            DeliveryMethod = SmtpDeliveryMethod.SpecifiedPickupDirectory,
                            PickupDirectoryLocation = pickupDir
                        };
                        pickupClient.Send(message);
                    }

                    Console.WriteLine("Email with inline SVG prepared successfully.");
                }
                catch (Exception sendEx)
                {
                    Console.WriteLine($"Failed to write email file: {sendEx.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
