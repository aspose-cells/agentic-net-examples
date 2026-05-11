using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using Aspose.Cells;

public class MergeAndEmailDemo
{
    public static void Run()
    {
        // Create the source workbook and add sample data
        Workbook sourceWorkbook = new Workbook();
        sourceWorkbook.Worksheets[0].Cells["A1"].PutValue("Source Data");

        // Create the destination workbook (XLSX format) and add sample data
        Workbook destWorkbook = new Workbook(FileFormatType.Xlsx);
        destWorkbook.Worksheets[0].Cells["B2"].PutValue("Destination Data");

        // Combine the source workbook into the destination workbook
        destWorkbook.Combine(sourceWorkbook);

        // Define a temporary file path for the merged workbook
        string mergedFilePath = Path.Combine(Path.GetTempPath(), "MergedWorkbook.xlsx");

        // Save the combined workbook
        destWorkbook.Save(mergedFilePath, SaveFormat.Xlsx);

        // Prepare the email message
        using (MailMessage message = new MailMessage())
        {
            message.From = new MailAddress("sender@example.com");
            message.To.Add("recipient@example.com");
            message.Subject = "Merged Workbook Attachment";
            message.Body = "Please find the merged workbook attached.";

            // Attach the merged workbook file
            using (Attachment attachment = new Attachment(mergedFilePath))
            {
                message.Attachments.Add(attachment);

                // Configure the SMTP client to use a local pickup directory (no external server needed)
                using (SmtpClient smtp = new SmtpClient())
                {
                    smtp.DeliveryMethod = SmtpDeliveryMethod.SpecifiedPickupDirectory;
                    string pickupDir = Path.Combine(Path.GetTempPath(), "MailPickup");
                    Directory.CreateDirectory(pickupDir);
                    smtp.PickupDirectoryLocation = pickupDir;

                    // Send the email (it will be saved to the pickup directory)
                    smtp.Send(message);
                }
            }
        }

        // Clean up the temporary file
        if (File.Exists(mergedFilePath))
        {
            File.Delete(mergedFilePath);
        }

        // Dispose disposable objects
        sourceWorkbook.Dispose();
        destWorkbook.Dispose();
    }
}

public class Program
{
    public static void Main()
    {
        MergeAndEmailDemo.Run();
    }
}