// Title: Generate a TIFF image from an Excel worksheet and send it as an email attachment using Aspose.Cells for .NET
// Common Searches: how to export an Excel worksheet to TIFF and email it using C# Aspose.Cells | C# Aspose.Cells save first worksheet as TIFF then attach to MailMessage | send Excel sheet as a TIFF attachment via SMTP in .NET | create temporary TIFF file from workbook for email distribution with Aspose.Cells | Aspose.Cells convert worksheet to image and send through System.Net.Mail
// Tags: Aspose.Cells worksheet to TIFF conversion | C# attach file to MailMessage | SMTP email with attachment in .NET | temporary file cleanup after email send | export Excel as image for distribution

using System;
using System.IO;
using System.Net.Mail;
using Aspose.Cells;

// The example loads or creates an Excel workbook, uses Aspose.Cells to save the first worksheet as a TIFF image in a temporary location, builds a MailMessage, attaches the TIFF file, sends the email via an SMTP client, and finally deletes the temporary TIFF file while handling potential errors.
class Program
{
    static void Main()
    {
        try
        {
            // Define input workbook path
            string inputPath = @"C:\Input\Sample.xlsx";

            // Load existing workbook or create a new one if the file does not exist
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Cells["A1"].PutValue("Sample Data");
                try
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(inputPath));
                    workbook.Save(inputPath);
                }
                catch
                {
                    // Ignored – creation of the sample file is optional
                }
            }

            // Convert the first worksheet to a TIFF image saved in a temporary location
            string tiffPath = Path.Combine(Path.GetTempPath(), "WorksheetImage.tiff");
            workbook.Save(tiffPath, SaveFormat.Tiff);

            // Prepare the email message
            using (MailMessage mail = new MailMessage())
            {
                mail.From = new MailAddress("sender@example.com");
                mail.To.Add("recipient@example.com");
                mail.Subject = "Converted TIFF Image";
                mail.Body = "Please find the converted TIFF image attached.";

                // Attach the generated TIFF file
                if (File.Exists(tiffPath))
                {
                    mail.Attachments.Add(new Attachment(tiffPath));
                }

                // Configure the SMTP client (replace with valid server details)
                using (SmtpClient smtp = new SmtpClient("smtp.example.com", 587))
                {
                    smtp.Credentials = new System.Net.NetworkCredential("username", "password");
                    smtp.EnableSsl = true;

                    try
                    {
                        smtp.Send(mail);
                        Console.WriteLine("Email sent successfully.");
                    }
                    catch (SmtpException smtpEx)
                    {
                        Console.WriteLine($"SMTP error: {smtpEx.Message}");
                    }
                }
            }

            // Clean up the temporary TIFF file
            if (File.Exists(tiffPath))
            {
                try
                {
                    File.Delete(tiffPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to delete temporary file: {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
