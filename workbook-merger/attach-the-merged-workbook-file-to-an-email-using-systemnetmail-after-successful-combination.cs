// Title: Combine Excel Workbooks with Aspose.Cells and Email the Merged File via System.Net.Mail (C#)
// Description: C# sample that verifies two source Excel files, creates them if missing, merges the second workbook into the first using Aspose.Cells Workbook.Combine, saves the result, and sends it as an email attachment through System.Net.Mail with an SMTP client. Includes basic error handling for file I/O and SMTP transmission.
// Keywords: Aspose.Cells combine workbooks | C# merge Excel files | Workbook.Combine Aspose | System.Net.Mail attachment | send Excel via SMTP | C# email attachment example | automated Excel reporting | merge and email Excel | Aspose.Cells .NET | SMTP client C#
// Common Searches: how to merge two Excel files with Aspose.Cells C# | attach generated workbook to email using System.Net.Mail | send merged Excel workbook via SMTP in .NET | Aspose.Cells Workbook.Combine example | C# code to email an Excel file attachment
// Developer Intent: The developer needs to combine multiple Excel workbooks using Aspose.Cells and automatically deliver the merged file as an email attachment through System.Net.Mail.
// Use Cases: Daily consolidation of departmental spreadsheets and automatic distribution to management. | Batch processing of client‑uploaded Excel files, merging them into a single report, then emailing the result. | Workflow automation where a merged workbook triggers a notification email to stakeholders after successful processing.
// AI Prompts: Generate C# code that merges an arbitrary number of Excel workbooks with Aspose.Cells and attaches the result to an email with configurable SMTP settings. | Add robust error handling and retry logic to the email‑sending block, including detailed logging and proper disposal of resources. | Show how to stream the merged workbook directly into a MailMessage attachment without writing a temporary file to disk.

using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using Aspose.Cells;

namespace AsposeCellsEmailDemo
{
    // C# sample that verifies two source Excel files, creates them if missing, merges the second workbook into the first using Aspose.Cells Workbook.Combine, saves the result, and sends it as an email attachment through System.Net.Mail with an SMTP client. Includes basic error handling for file I/O and SMTP transmission.
    class Program
    {
        static void Main()
        {
            // Paths to the workbooks that need to be merged
            string firstWorkbookPath = "FirstWorkbook.xlsx";
            string secondWorkbookPath = "SecondWorkbook.xlsx";

            // Path for the merged workbook
            string mergedWorkbookPath = "MergedWorkbook.xlsx";

            try
            {
                // Ensure source workbooks exist; create empty ones if missing
                if (!File.Exists(firstWorkbookPath))
                {
                    new Workbook().Save(firstWorkbookPath, SaveFormat.Xlsx);
                }

                if (!File.Exists(secondWorkbookPath))
                {
                    new Workbook().Save(secondWorkbookPath, SaveFormat.Xlsx);
                }

                // Load the first workbook (destination)
                Workbook destWorkbook = new Workbook(firstWorkbookPath);

                // Load the second workbook (source)
                Workbook sourceWorkbook = new Workbook(secondWorkbookPath);

                // Combine the source workbook into the destination workbook
                destWorkbook.Combine(sourceWorkbook);

                // Save the merged workbook to disk
                destWorkbook.Save(mergedWorkbookPath, SaveFormat.Xlsx);

                // Prepare the email message
                using (MailMessage mail = new MailMessage())
                {
                    mail.From = new MailAddress("sender@example.com");
                    mail.To.Add("recipient@example.com");
                    mail.Subject = "Merged Workbook Attachment";
                    mail.Body = "Please find the merged workbook attached.";

                    // Attach the merged workbook file
                    using (Attachment attachment = new Attachment(mergedWorkbookPath))
                    {
                        mail.Attachments.Add(attachment);

                        // Configure the SMTP client (replace with actual SMTP server details)
                        using (SmtpClient smtp = new SmtpClient("smtp.example.com", 587))
                        {
                            smtp.Credentials = new NetworkCredential("smtp_user", "smtp_password");
                            smtp.EnableSsl = true;

                            try
                            {
                                // Send the email
                                smtp.Send(mail);
                                Console.WriteLine("Email sent with merged workbook attached.");
                            }
                            catch (SmtpException smtpEx)
                            {
                                Console.WriteLine($"Failed to send email: {smtpEx.Message}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
