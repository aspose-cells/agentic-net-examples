// Title: Attach a converted TIFF file to an email and send it with C# System.Net.Mail and SmtpClient
// AI Prompts: Generate a C# console program that checks for the existence of a .tif file, creates a MailMessage, adds the file as an Attachment with MIME type image/tiff, and sends it through an SSL‑enabled SmtpClient. | Modify the email‑sending code to log detailed errors when the TIFF file is missing or the SMTP transmission fails, and ensure all disposable objects are correctly disposed. | Extend the example to read SMTP configuration (host, port, credentials, SSL) from an appsettings.json file and attach multiple TIFF files from a folder.
// Common Searches: C# code to email a .tif image using System.Net.Mail | how to set MIME type image/tiff for an email attachment in .NET | sending multiple TIFF files as email attachments with SmtpClient in C# | validate file path before attaching image to MailMessage in C# | configure SSL SMTP client for sending image attachments in .NET Core
// Tags: C# MailMessage attach TIFF | System.Net.Mail set image/tiff MIME | SmtpClient SSL send image attachment | validate TIFF file existence before email | read SMTP settings from appsettings.json C#

using System;
using System.Net;
using System.Net.Mail;
using System.IO;

// The example verifies that a converted TIFF file exists, creates a MailMessage with sender, recipient, subject, and body, attaches the TIFF file while explicitly setting its MIME type to image/tiff, configures an SSL‑enabled SmtpClient with credentials, sends the email, handles missing‑file and transmission errors, and disposes all resources. Optional enhancements include loading SMTP settings from configuration and attaching multiple TIFF files.
class EmailSender
{
    static void Main()
    {
        // Path to the TIFF image that was created earlier
        string tiffFilePath = @"C:\Images\ConvertedImage.tif";

        // Verify that the file exists before attempting to attach it
        if (!File.Exists(tiffFilePath))
        {
            Console.WriteLine("TIFF file not found: " + tiffFilePath);
            return;
        }

        // Create a new mail message
        MailMessage message = new MailMessage();

        // Set sender and recipient addresses
        message.From = new MailAddress("sender@example.com");
        message.To.Add(new MailAddress("recipient@example.com"));

        // Set subject and body
        message.Subject = "Converted TIFF Image Attachment";
        message.Body = "Please find the converted TIFF image attached.";

        // Attach the TIFF file
        Attachment tiffAttachment = new Attachment(tiffFilePath);
        // Optionally set the MIME type explicitly
        tiffAttachment.ContentType.MediaType = "image/tiff";
        message.Attachments.Add(tiffAttachment);

        // Configure the SMTP client (adjust host, port, and credentials as needed)
        SmtpClient smtpClient = new SmtpClient("smtp.example.com", 587)
        {
            EnableSsl = true,
            Credentials = new NetworkCredential("smtp_user", "smtp_password")
        };

        try
        {
            // Send the email
            smtpClient.Send(message);
            Console.WriteLine("Email sent successfully with TIFF attachment.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error sending email: " + ex.Message);
        }
        finally
        {
            // Clean up resources
            tiffAttachment.Dispose();
            message.Dispose();
            smtpClient.Dispose();
        }
    }
}
