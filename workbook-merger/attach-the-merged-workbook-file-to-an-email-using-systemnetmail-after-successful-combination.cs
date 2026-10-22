// Title: Send a merged Excel workbook created with Aspose.Cells as an email attachment using System.Net.Mail (C#)
// AI Prompts: Write a C# method that accepts the path of an .xlsx workbook and sends it as an email attachment via SMTP, handling SSL and credential setup with System.Net.Mail. | Enhance the email routine to attach several Aspose.Cells‑generated workbooks, allow an optional HTML body, and reuse the existing SMTP configuration.
// Common Searches: C# how to attach an Aspose.Cells generated .xlsx file to an email with SmtpClient | sending merged workbook via System.Net.Mail with SSL authentication | example of emailing Excel file after combining worksheets in Aspose.Cells | C# code to email multiple Excel workbooks using System.Net.Mail | attach generated Excel file to email and handle missing file exception in C#
// Tags: SMTP email with Excel attachment | Aspose.Cells merged workbook email | C# send .xlsx via SmtpClient | handle missing attachment file C# | email multiple workbooks using SmtpClient

using Aspose.Cells;
using System;
using System.IO;
using System.Net;
using System.Net.Mail;

// The sample ensures a merged workbook exists (creating a simple one if needed), configures SMTP credentials, and uses System.Net.Mail to compose and send an email with the workbook attached, including file‑existence validation and SSL‑enabled authentication.
class Program
{
    static void Main()
    {
        // Path to the merged workbook that was created earlier
        string mergedFilePath = "MergedWorkbook.xlsx";

        // Ensure the workbook exists; create a simple one if it does not
        try
        {
            if (!File.Exists(mergedFilePath))
            {
                var wb = new Workbook();
                wb.Worksheets[0].Cells["A1"].PutValue("Sample data");
                wb.Save(mergedFilePath);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error preparing workbook: {ex.Message}");
            return;
        }

        // SMTP server configuration (replace with valid values for actual use)
        string smtpHost = "smtp.example.com";
        int smtpPort = 587;
        string smtpUser = "user@example.com";
        string smtpPass = "password";

        // Email details
        string fromAddress = "user@example.com";
        string toAddress = "recipient@example.com";
        string subject = "Merged Workbook";
        string body = "Please find the merged workbook attached.";

        // Send the email with the merged workbook attached
        try
        {
            SendEmailWithAttachment(
                smtpHost, smtpPort, smtpUser, smtpPass,
                fromAddress, toAddress, subject, body,
                mergedFilePath);
            Console.WriteLine("Email sent successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error sending email: {ex.Message}");
        }
    }

    static void SendEmailWithAttachment(
        string host, int port, string username, string password,
        string from, string to, string subject, string body,
        string attachmentPath)
    {
        // Verify attachment file exists before creating the MailMessage
        if (!File.Exists(attachmentPath))
            throw new FileNotFoundException("Attachment file not found.", attachmentPath);

        using (MailMessage mail = new MailMessage())
        {
            mail.From = new MailAddress(from);
            mail.To.Add(to);
            mail.Subject = subject;
            mail.Body = body;
            mail.IsBodyHtml = false;

            // Attach the merged workbook file
            mail.Attachments.Add(new Attachment(attachmentPath));

            using (SmtpClient smtp = new SmtpClient(host, port))
            {
                smtp.Credentials = new NetworkCredential(username, password);
                smtp.EnableSsl = true; // Set to false if SSL is not required
                smtp.Send(mail);
            }
        }
    }
}
