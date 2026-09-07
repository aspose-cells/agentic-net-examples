// Title: Generate HTML from an Aspose.Cells workbook and send it as the body of a System.Net.Mail newsletter in C#
// AI Prompts: Write a C# method that creates an Aspose.Cells workbook, saves it as HTML in a MemoryStream, and uses SmtpClient to deliver the HTML as the email body. | Extend the program to also attach the generated HTML file while keeping the same content as the message body. | Add custom inline CSS to the HTML produced by Aspose.Cells before sending the email with System.Net.Mail.
// Common Searches: how to convert Aspose.Cells workbook to html and email it using System.Net.Mail c# | c# send newsletter with html generated from Excel workbook via smtp | aspocells save as html stream and embed in email body .net | using memory stream to send html email from Aspose.Cells workbook c#
// Tags: Aspose.Cells workbook to HTML conversion | System.Net.Mail HTML email body | C# send newsletter via SMTP with generated HTML | MemoryStream HTML generation from Excel | inline CSS injection into Aspose.Cells HTML output

using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using Aspose.Cells;

// The example creates an Aspose.Cells workbook, populates it with data, converts it to HTML stored in a MemoryStream, reads the HTML string, and sends it as the body of an email using System.Net.Mail with configurable SMTP settings.
class NewsletterSender
{
    static void Main()
    {
        try
        {
            // Create a new workbook (using Aspose.Cells)
            Workbook workbook = new Workbook();

            // Populate the first worksheet with sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Welcome to our Newsletter");
            sheet.Cells["A2"].PutValue(DateTime.Now.ToString("D"));
            sheet.Cells["A4"].PutValue("Here is some important information:");
            sheet.Cells["A5"].PutValue("• Item 1");
            sheet.Cells["A6"].PutValue("• Item 2");
            sheet.Cells["A7"].PutValue("• Item 3");

            // Convert the workbook to HTML and store it in a memory stream
            using (MemoryStream htmlStream = new MemoryStream())
            {
                workbook.Save(htmlStream, SaveFormat.Html);
                htmlStream.Position = 0; // Reset stream position for reading

                // Read the generated HTML into a string
                string htmlBody = new StreamReader(htmlStream).ReadToEnd();

                // Create the email message
                using (MailMessage mail = new MailMessage())
                {
                    mail.From = new MailAddress("sender@example.com");          // Replace with actual sender
                    mail.To.Add("recipient@example.com");                      // Replace with actual recipient
                    mail.Subject = "Monthly Newsletter";
                    mail.Body = htmlBody;                                      // Embed the HTML as the email body
                    mail.IsBodyHtml = true;                                    // Indicate that the body is HTML

                    // Configure the SMTP client (adjust settings as needed)
                    using (SmtpClient smtp = new SmtpClient("smtp.example.com")) // Replace with actual SMTP server
                    {
                        smtp.Port = 587;                                           // Common port for TLS
                        smtp.Credentials = new NetworkCredential("username", "password"); // Replace with credentials
                        smtp.EnableSsl = true;                                     // Use SSL/TLS

                        try
                        {
                            // Send the email
                            smtp.Send(mail);
                            Console.WriteLine("Email sent successfully.");
                        }
                        catch (SmtpException smtpEx)
                        {
                            Console.WriteLine($"SMTP error: {smtpEx.Message}");
                            // Optionally log or handle the error further
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
