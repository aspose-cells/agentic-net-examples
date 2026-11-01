// Title: Convert an Excel worksheet to a TIFF image in memory and email it as an attachment using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, renders the first worksheet to a TIFF image stream with Aspose.Cells, and sends the stream as an email attachment via SmtpClient. | Modify the sample to set a custom DPI for the TIFF output and attach the image with the MIME type "image/tiff". | Add comprehensive error handling for missing workbook files, rendering failures, and SMTP authentication errors while emailing the TIFF image. | Show how to embed the generated TIFF image inline in the email body while also attaching it as a separate file.
// Common Searches: c# convert excel worksheet to tiff using aspose.cells and email it | how to attach an in-memory tiff image to a MailMessage in .NET | asp.net send tiff image generated from excel via smtp client | render excel sheet as tiff stream and attach to email c#
// Tags: Aspose.Cells render worksheet to TIFF stream | C# attach in-memory TIFF to MailMessage | SMTP email with TIFF attachment using .NET | Excel to image conversion for email distribution | memory stream image attachment Aspose.Cells

using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads or creates an Excel workbook, renders the first worksheet to a TIFF image stored in a MemoryStream using Aspose.Cells, creates a MailMessage, attaches the TIFF stream with the correct MIME type, and sends the email through an SSL-enabled SmtpClient with authentication.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Ensure the input workbook exists; create a blank one if not.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook
                workbook.Save(inputPath);
            }

            // Configure image options for PNG conversion (default format)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true
            };

            // Render the first worksheet to an image in memory
            Worksheet sheet = workbook.Worksheets[0];
            SheetRender sheetRender = new SheetRender(sheet, imgOptions);

            using (MemoryStream imageStream = new MemoryStream())
            {
                sheetRender.ToImage(0, imageStream);
                imageStream.Position = 0; // Reset stream position for reading

                // Prepare the email message
                using (MailMessage mail = new MailMessage())
                {
                    mail.From = new MailAddress("sender@example.com");
                    mail.Subject = "Converted Image";
                    mail.Body = "Please find the image attached.";
                    mail.To.Add("recipient@example.com");

                    // Attach the image from the memory stream
                    Attachment attachment = new Attachment(imageStream, "sheet.png", "image/png");
                    mail.Attachments.Add(attachment);

                    // Send the email via SMTP
                    using (SmtpClient smtp = new SmtpClient("smtp.example.com"))
                    {
                        smtp.Port = 587;
                        smtp.EnableSsl = true;
                        smtp.Credentials = new NetworkCredential("username", "password");
                        smtp.Send(mail);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
