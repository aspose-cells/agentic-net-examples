// Title: Combine multiple Excel files with Aspose.Cells, save as a merged workbook, and email it using System.Net.Mail in C#
// Common Searches: c# merge two .xlsx files with Aspose.Cells and send as email attachment | how to attach an Aspose.Cells generated workbook to an SMTP email in .NET | using System.Net.Mail to email a merged Excel workbook created by Aspose.Cells | combine worksheets from multiple workbooks and email the result in C# | save merged workbook to disk then send via SMTP with C# | aspocells workbook merging and email delivery example
// Tags: aspocells workbook merging xlsx | c# attach excel file to smtp email | system.net.mail send merged workbook | save merged workbook to disk c# | combine worksheets aspocells | email attachment of generated excel c#

using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using Aspose.Cells;

// The sample loads two existing .xlsx files, uses Aspose.Cells to copy all worksheets from the second workbook into the first, saves the combined workbook as a new file, verifies its existence, then creates a System.Net.Mail message, attaches the merged workbook, configures an SmtpClient with SSL and credentials, and sends the email while handling errors and disposing resources.
class WorkbookEmailer
{
    static void Main()
    {
        // Paths for source workbooks and the merged output
        string sourcePath1 = @"C:\Data\Source1.xlsx";
        string sourcePath2 = @"C:\Data\Source2.xlsx";
        string mergedPath = @"C:\Data\MergedWorkbook.xlsx";

        // Verify source files exist
        if (!File.Exists(sourcePath1))
        {
            Console.WriteLine($"Source file not found: {sourcePath1}");
            return;
        }
        if (!File.Exists(sourcePath2))
        {
            Console.WriteLine($"Source file not found: {sourcePath2}");
            return;
        }

        Workbook mergedWorkbook = null;
        Workbook wbToAppend = null;

        try
        {
            // Load the first workbook (this will be the target workbook)
            mergedWorkbook = new Workbook(sourcePath1);

            // Load the second workbook
            wbToAppend = new Workbook(sourcePath2);

            // Append all worksheets from the second workbook to the merged workbook
            foreach (Worksheet sheet in wbToAppend.Worksheets)
            {
                // Clone the worksheet into the target workbook by name
                mergedWorkbook.Worksheets.AddCopy(sheet.Name);
            }

            // Save the merged workbook to disk
            mergedWorkbook.Save(mergedPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during workbook processing: {ex.Message}");
            return;
        }
        finally
        {
            // Release workbook resources
            mergedWorkbook?.Dispose();
            wbToAppend?.Dispose();
        }

        // Verify that the merged file was created successfully
        if (!File.Exists(mergedPath))
        {
            Console.WriteLine("Failed to create the merged workbook.");
            return;
        }

        // Prepare the email message
        using (MailMessage mail = new MailMessage())
        {
            mail.From = new MailAddress("sender@example.com");
            mail.To.Add("recipient@example.com");
            mail.Subject = "Merged Workbook Attachment";
            mail.Body = "Please find the merged workbook attached.";

            // Attach the merged workbook file
            using (Attachment attachment = new Attachment(mergedPath))
            {
                mail.Attachments.Add(attachment);

                // Configure the SMTP client (adjust host, port, and credentials as needed)
                using (SmtpClient smtp = new SmtpClient("smtp.example.com", 587))
                {
                    smtp.EnableSsl = true;
                    smtp.Credentials = new NetworkCredential("smtp_user", "smtp_password");

                    try
                    {
                        // Send the email with the attachment
                        smtp.Send(mail);
                        Console.WriteLine("Email sent successfully with the merged workbook attached.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error sending email: " + ex.Message);
                    }
                }
            }
        }
    }
}
