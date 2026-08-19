// Title: Combine Excel Workbooks with Aspose.Cells and Email as Attachment (C#)
// Description: Load two workbooks, merge the source into the destination using Aspose.Cells' Combine method, save the result to a temporary file, attach it to a System.Net.Mail message, send via an SSL‑enabled SmtpClient, and delete the temporary file after delivery. Includes basic error handling for missing files and SMTP failures.
// Keywords: Aspose.Cells combine workbooks C# | merge Excel files Aspose.Cells | email Excel attachment System.Net.Mail | C# send merged workbook via SMTP | temporary file cleanup after email | SMTP SSL C# Aspose.Cells example
// Common Searches: how to merge two Excel files with Aspose.Cells and email them | C# combine workbooks and send as attachment using System.Net.Mail | Aspose.Cells Combine method example with SMTP | delete temporary Excel file after sending email in .NET | send merged workbook via Gmail SMTP C#
// Developer Intent: The developer needs to programmatically merge two Excel workbooks with Aspose.Cells, attach the merged file to an email, and deliver it through an SMTP server.
// Use Cases: Automated daily reporting: merge a template and a data workbook, then email the combined report to stakeholders. | Consolidation service: combine departmental spreadsheets into a single file and distribute it to a mailing list. | Web API endpoint: accept uploaded Excel files, merge them on the server, and return the combined workbook as an email attachment.
// AI Prompts: Create C# code that uses Aspose.Cells to combine two workbooks, saves the result to a temporary file, attaches it to a MailMessage, sends via SmtpClient with SSL, and removes the temporary file afterward. | Add robust error handling for file not found, SMTP authentication failures, and general exceptions in the workbook merge‑and‑email workflow. | Show how to read SMTP host, port, username, and password from appsettings.json and inject them into the SmtpClient configuration for sending the merged workbook.

using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using Aspose.Cells;

namespace AsposeCellsEmailDemo
{
    // Load two workbooks, merge the source into the destination using Aspose.Cells' Combine method, save the result to a temporary file, attach it to a System.Net.Mail message, send via an SSL‑enabled SmtpClient, and delete the temporary file after delivery. Includes basic error handling for missing files and SMTP failures.
    public class MergeAndEmail
    {
        public static void Run()
        {
            try
            {
                // Paths for the source and destination workbooks
                string sourcePath = "Source.xlsx";
                string destinationPath = "Destination.xlsx";

                // Verify source file exists
                if (!File.Exists(sourcePath))
                    throw new FileNotFoundException($"Source file not found: {sourcePath}");

                // Verify destination file exists
                if (!File.Exists(destinationPath))
                    throw new FileNotFoundException($"Destination file not found: {destinationPath}");

                // Load the source workbook from file
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Load the destination workbook from file
                Workbook destinationWorkbook = new Workbook(destinationPath);

                // Combine the source workbook into the destination workbook
                destinationWorkbook.Combine(sourceWorkbook);

                // Save the combined workbook to a temporary file
                string combinedPath = "CombinedWorkbook.xlsx";
                destinationWorkbook.Save(combinedPath, SaveFormat.Xlsx);

                // Prepare the email message
                using (MailMessage mail = new MailMessage())
                {
                    mail.From = new MailAddress("sender@example.com");
                    mail.To.Add("recipient@example.com");
                    mail.Subject = "Combined Workbook";
                    mail.Body = "Please find the combined workbook attached.";

                    // Attach the combined workbook file
                    Attachment attachment = new Attachment(combinedPath);
                    mail.Attachments.Add(attachment);

                    // Configure the SMTP client (replace with real server details)
                    using (SmtpClient smtp = new SmtpClient("smtp.example.com", 587))
                    {
                        smtp.Credentials = new NetworkCredential("username", "password");
                        smtp.EnableSsl = true;

                        // Send the email
                        smtp.Send(mail);
                    }
                }

                // Clean up the temporary combined file
                if (File.Exists(combinedPath))
                {
                    File.Delete(combinedPath);
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
                // Optionally rethrow or handle specific exceptions
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            MergeAndEmail.Run();
        }
    }
}
