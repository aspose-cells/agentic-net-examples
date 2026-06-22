using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfAttachmentDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("PDF with Embedded Image Attachment");

            // Path to the image that will be embedded as an attachment
            string imagePath = "sample.png";

            // Ensure the image file exists; for demo purposes create a simple PNG if missing
            if (!File.Exists(imagePath))
            {
                // Create a 1x1 pixel PNG (binary content)
                byte[] pngBytes = Convert.FromBase64String(
                    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK2cAAAAASUVORK5CYII=");
                File.WriteAllBytes(imagePath, pngBytes);
            }

            // Add the image as an OLE object (attachment) to the worksheet
            // Parameters: row, column, width, height, byte[] of the file
            int oleIndex = sheet.OleObjects.Add(5, 0, 200, 200, File.ReadAllBytes(imagePath));
            // Specify the file format type of the embedded object
            sheet.OleObjects[oleIndex].FileFormatType = FileFormatType.Png;
            // Optionally display the attachment as an icon
            sheet.OleObjects[oleIndex].DisplayAsIcon = true;

            // Configure PDF save options to embed attachments
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                EmbedAttachments = true   // Enable embedding of OLE attachments
            };

            // Save the workbook as PDF with the embedded image attachment
            string pdfPath = "WorkbookWithImageAttachment.pdf";
            workbook.Save(pdfPath, pdfOptions);

            // Clean up the temporary image file if it was created by this demo
            if (File.Exists(imagePath))
            {
                File.Delete(imagePath);
            }

            Console.WriteLine($"PDF saved to '{pdfPath}' with the image embedded as an attachment.");
        }
    }
}