using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("PDF with Embedded Image Attachment");

        // Path to the image file that will be embedded
        string imagePath = "sample.png";

        // Ensure the image file exists (in a real scenario provide a valid image)
        if (!File.Exists(imagePath))
        {
            // Create a placeholder empty file for demonstration purposes
            File.WriteAllBytes(imagePath, new byte[0]);
        }

        // Read the image bytes
        byte[] imageBytes = File.ReadAllBytes(imagePath);

        // Add the image as an OLE object to the worksheet
        int oleIndex = sheet.OleObjects.Add(5, 5, 200, 200, imageBytes);
        OleObject oleObject = sheet.OleObjects[oleIndex];
        oleObject.FileFormatType = FileFormatType.Png;   // Specify the image format
        oleObject.DisplayAsIcon = true;                  // Optional: display as an icon

        // Configure PDF save options to embed OLE attachments
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            EmbedAttachments = true
        };

        // Save the workbook as a PDF with the embedded image attachment
        workbook.Save("PdfWithEmbeddedImage.pdf", pdfOptions);

        // Optional cleanup of the temporary image file
        // File.Delete(imagePath);
    }
}

// Author: Example demonstrating how to embed an image as a PDF attachment using Aspose.Cells.