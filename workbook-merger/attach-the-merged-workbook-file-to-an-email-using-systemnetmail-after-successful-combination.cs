// Title: Merge two Excel workbooks with Aspose.Cells and email the combined file using System.Net.Mail in C#
// AI Prompts: Use Aspose.Cells to combine two .xlsx workbooks, save the merged workbook, and attach it to an email with SmtpClient in C#. | Create missing source Excel files, merge them using Workbook.Combine, then send the resulting file as an attachment via System.Net.Mail. | Write C# code that loads two workbooks, merges them, saves as Xlsx, and delivers the file through an SMTP server.
// Common Searches: c# how to combine multiple Excel files with Aspose.Cells and email the result | asp.net send merged workbook as attachment using System.Net.Mail | example of Workbook.Combine followed by SmtpClient send in C# | merge two .xlsx files and attach to email programmatically
// Tags: Aspose.Cells combine workbooks Xlsx | C# send merged Excel via SmtpClient | Workbook.Combine method example | System.Net.Mail attachment of Excel file | programmatic Excel merge and email

using System;
using System.IO;
using System.Net.Mail;
using Aspose.Cells;

// The sample ensures two source .xlsx files exist (creating simple workbooks if needed), merges the second workbook into the first with Aspose.Cells' Combine method, saves the merged result as MergedWorkbook.xlsx, then composes an email using System.Net.Mail, attaches the merged file, and sends it through an SMTP client.
class Program
{
    static void Main()
    {
        // Paths to the source workbooks and the merged output file
        string sourcePath1 = "Source1.xlsx";
        string sourcePath2 = "Source2.xlsx";
        string mergedPath = "MergedWorkbook.xlsx";

        try
        {
            // Ensure source files exist; create simple workbooks if they are missing
            if (!File.Exists(sourcePath1))
            {
                var wb1 = new Workbook();
                wb1.Worksheets[0].Cells["A1"].PutValue("Data from Source1");
                wb1.Save(sourcePath1);
                wb1.Dispose();
            }

            if (!File.Exists(sourcePath2))
            {
                var wb2 = new Workbook();
                wb2.Worksheets[0].Cells["A1"].PutValue("Data from Source2");
                wb2.Save(sourcePath2);
                wb2.Dispose();
            }

            // Load the first workbook (will become the destination workbook)
            using (var destWorkbook = new Workbook(sourcePath1))
            {
                // Load the second workbook to be merged
                using (var secondWorkbook = new Workbook(sourcePath2))
                {
                    // Combine the second workbook into the destination workbook
                    destWorkbook.Combine(secondWorkbook);

                    // Save the combined workbook to disk
                    destWorkbook.Save(mergedPath, SaveFormat.Xlsx);
                }
            }

            // Create and send email with the merged workbook attached
            using (var mail = new MailMessage())
            {
                mail.From = new MailAddress("sender@example.com");
                mail.To.Add("recipient@example.com");
                mail.Subject = "Merged Workbook Attachment";
                mail.Body = "The merged workbook is attached.";

                // Attach the merged workbook file
                using (var attachment = new Attachment(mergedPath))
                {
                    mail.Attachments.Add(attachment);

                    // Configure the SMTP client (replace with actual server details)
                    using (var smtp = new SmtpClient("smtp.example.com"))
                    {
                        smtp.Port = 587;
                        smtp.Credentials = new System.Net.NetworkCredential("username", "password");
                        smtp.EnableSsl = true;

                        // Send the email
                        smtp.Send(mail);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // Log or display the error as needed
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
