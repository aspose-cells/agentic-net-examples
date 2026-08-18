// Title: C# – Render Excel worksheet to TIFF and attach to email with Aspose.Cells
// Description: Shows how to create a workbook, render the first worksheet to a single‑page TIFF using Aspose.Cells, load it into a MemoryStream, and add it as an image/tiff attachment to a System.Net.Mail MailMessage ready for SMTP sending.
// Keywords: Aspose.Cells | C# TIFF rendering | Excel to TIFF | MailMessage attachment | MemoryStream email | SheetRender | ImageOrPrintOptions | System.Net.Mail | SMTP attachment | no temporary file
// Common Searches: convert Excel worksheet to TIFF in C# | attach generated TIFF to email without saving file | Aspose.Cells render sheet as TIFF stream | C# send Excel snapshot as image attachment | System.Net.Mail attach MemoryStream
// Developer Intent: Create a TIFF image of an Excel sheet and embed it directly in an email message without creating a physical file.
// Use Cases: Automated reporting: email a one‑page TIFF of a financial summary to stakeholders. | Alert system: send a spreadsheet snapshot as an image attachment in notification emails. | Compliance archive: deliver Excel data as a non‑editable TIFF via SMTP.
// AI Prompts: Generate C# code that uses Aspose.Cells to render a worksheet to a TIFF MemoryStream and attaches it to a MailMessage. | Provide an example of emailing an Excel sheet as a TIFF image without writing the file to disk, using Aspose.Cells and System.Net.Mail. | Explain how to configure ImageOrPrintOptions for a single‑page TIFF, render with SheetRender, reset the stream, and create an Attachment for SMTP delivery.

using System;
using System.IO;
using System.Net.Mail;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// Shows how to create a workbook, render the first worksheet to a single‑page TIFF using Aspose.Cells, load it into a MemoryStream, and add it as an image/tiff attachment to a System.Net.Mail MailMessage ready for SMTP sending.
class Program
{
    static void Main()
    {
        // Create a new workbook and access the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample content to the worksheet
        worksheet.Cells["A1"].PutValue("Aspose.Cells TIFF Email Demo");

        // Configure image rendering options for TIFF
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            OnePagePerSheet = true,
            ImageType = ImageType.Tiff
        };

        // Render the worksheet to a TIFF image using a memory stream
        using (MemoryStream tiffStream = new MemoryStream())
        {
            SheetRender renderer = new SheetRender(worksheet, options);
            renderer.ToTiff(tiffStream); // Uses the provided ToTiff(Stream) rule

            // Reset the stream position before reading
            tiffStream.Position = 0;

            // Create an email message
            MailMessage mail = new MailMessage
            {
                From = new MailAddress("sender@example.com"),
                Subject = "Worksheet as TIFF attachment",
                Body = "Please find the attached TIFF image of the worksheet."
            };
            mail.To.Add("recipient@example.com");

            // Attach the TIFF image from the memory stream
            Attachment attachment = new Attachment(tiffStream, "Worksheet.tiff", "image/tiff");
            mail.Attachments.Add(attachment);

            // Optional: send the email using an SMTP client (configure as needed)
            // SmtpClient client = new SmtpClient("smtp.example.com");
            // client.Credentials = new System.Net.NetworkCredential("username", "password");
            // client.Send(mail);
        }
    }
}
