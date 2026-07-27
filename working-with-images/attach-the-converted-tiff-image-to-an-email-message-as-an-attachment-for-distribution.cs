using System;
using System.IO;
using System.Net.Mail;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

class TiffEmailAttachmentDemo
{
    static void Main()
    {
        // Load or create a workbook
        Workbook workbook = new Workbook(); // create new workbook
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data to the worksheet
        sheet.Cells["A1"].PutValue("Aspose.Cells TIFF Email Demo");
        sheet.Cells["A2"].PutValue(DateTime.Now);

        // Configure image rendering options for TIFF
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = ImageType.Tiff,
            OnePagePerSheet = true
        };

        // Render the worksheet to a TIFF image in a memory stream
        using (MemoryStream tiffStream = new MemoryStream())
        {
            SheetRender renderer = new SheetRender(sheet, options);
            renderer.ToTiff(tiffStream);          // use provided ToTiff(Stream) rule
            tiffStream.Position = 0;              // reset stream for reading

            // Prepare the email message
            MailMessage message = new MailMessage();
            message.From = new MailAddress("sender@example.com");
            message.To.Add("recipient@example.com");
            message.Subject = "Worksheet as TIFF Attachment";
            message.Body = "Please find the worksheet rendered as a TIFF image attached.";

            // Create attachment from the TIFF stream
            Attachment tiffAttachment = new Attachment(tiffStream, "Worksheet.tiff", "image/tiff");
            message.Attachments.Add(tiffAttachment);

            // Send the email (SMTP settings must be configured appropriately)
            using (SmtpClient smtp = new SmtpClient("smtp.example.com", 587))
            {
                smtp.Credentials = new System.Net.NetworkCredential("username", "password");
                smtp.EnableSsl = true;

                // Uncomment the line below to actually send the email
                // smtp.Send(message);
            }

            // Dispose attachment (does not close the underlying stream because it's used by MailMessage)
            tiffAttachment.Dispose();
        }
    }
}