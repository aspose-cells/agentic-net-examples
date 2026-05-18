using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using Aspose.Cells;

class MergeAndEmail
{
    static void Main()
    {
        try
        {
            // Paths of the workbooks to be merged
            string sourcePath1 = "Source1.xlsx";
            string sourcePath2 = "Source2.xlsx";

            // Verify source files exist
            if (!File.Exists(sourcePath1))
                throw new FileNotFoundException($"Source file not found: {sourcePath1}");
            if (!File.Exists(sourcePath2))
                throw new FileNotFoundException($"Source file not found: {sourcePath2}");

            // Path for the merged workbook
            string mergedPath = "MergedWorkbook.xlsx";

            // Load the source workbooks
            using (Workbook sourceWorkbook1 = new Workbook(sourcePath1))
            using (Workbook sourceWorkbook2 = new Workbook(sourcePath2))
            // Create an empty destination workbook
            using (Workbook destWorkbook = new Workbook())
            {
                // Combine the source workbooks into the destination workbook
                destWorkbook.Combine(sourceWorkbook1);
                destWorkbook.Combine(sourceWorkbook2);

                // Save the merged workbook
                destWorkbook.Save(mergedPath, SaveFormat.Xlsx);
            }

            // Email configuration (replace with valid values)
            string smtpHost = "smtp.example.com";
            int smtpPort = 587;
            string smtpUser = "user@example.com";
            string smtpPass = "password";

            string fromAddress = "user@example.com";
            string toAddress = "recipient@example.com";
            string subject = "Merged Workbook";
            string body = "Please find the merged workbook attached.";

            // Create the email message
            using (MailMessage mail = new MailMessage())
            {
                mail.From = new MailAddress(fromAddress);
                mail.To.Add(toAddress);
                mail.Subject = subject;
                mail.Body = body;

                // Attach the merged workbook file
                if (!File.Exists(mergedPath))
                    throw new FileNotFoundException($"Merged file not found: {mergedPath}");
                mail.Attachments.Add(new Attachment(mergedPath));

                // Send the email
                using (SmtpClient client = new SmtpClient(smtpHost, smtpPort))
                {
                    client.EnableSsl = true;
                    client.Credentials = new NetworkCredential(smtpUser, smtpPass);
                    client.Send(mail);
                }
            }
        }
        catch (FileNotFoundException fnfEx)
        {
            Console.Error.WriteLine($"File error: {fnfEx.Message}");
        }
        catch (SmtpException smtpEx)
        {
            Console.Error.WriteLine($"SMTP error: {smtpEx.Message}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}