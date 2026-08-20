// Title: Embed an Image as a PDF Attachment with Aspose.Cells PdfSaveOptions (C#)
// Description: The sample creates a workbook, inserts a PNG file as an OLE object, configures it to show as an icon, enables the EmbedAttachments flag on PdfSaveOptions, and saves the workbook to PDF so the image is stored as an embedded attachment.
// Keywords: Aspose.Cells PDF attachment C# | PdfSaveOptions EmbedAttachments | add OLE image Aspose.Cells | export Excel to PDF with hidden files | C# embed PNG in PDF
// Common Searches: add PNG as attachment when saving Excel to PDF with Aspose.Cells | Aspose.Cells PdfSaveOptions EmbedAttachments example C# | C# embed OLE object and export to PDF | save workbook with embedded files using Aspose.Cells
// Developer Intent: Include an image file as an embedded attachment inside a PDF generated from an Excel workbook.
// Use Cases: Generate a PDF report that carries supporting diagrams as hidden attachments for later extraction. | Create an invoice PDF where product photos are attached but not displayed on the page. | Distribute technical documentation with reference images attached without cluttering the visible layout.
// AI Prompts: Show how to embed multiple images as PDF attachments using Aspose.Cells PdfSaveOptions in C#. | Explain how to extract embedded attachments from a PDF produced by Aspose.Cells. | Provide a version of the code that uses a MemoryStream for the image instead of a temporary file.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfAttachmentDemo
{
    // The sample creates a workbook, inserts a PNG file as an OLE object, configures it to show as an icon, enables the EmbedAttachments flag on PdfSaveOptions, and saves the workbook to PDF so the image is stored as an embedded attachment.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("PDF with Embedded Image Attachment");

            // Path to the image that will be embedded as an attachment
            string imagePath = "sampleImage.png";

            // Ensure the image file exists (for demo purposes we create a simple placeholder)
            if (!File.Exists(imagePath))
            {
                // Create a tiny PNG file (1x1 pixel) if it does not exist
                byte[] pngBytes = Convert.FromBase64String(
                    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=");
                File.WriteAllBytes(imagePath, pngBytes);
            }

            // Add the image as an OLE object (attachment) to the worksheet
            // Parameters: upper left row, upper left column, height, width, byte[] of the file
            int oleIndex = sheet.OleObjects.Add(5, 5, 100, 100, File.ReadAllBytes(imagePath));
            // Specify the file format type for the OLE object (PNG image)
            sheet.OleObjects[oleIndex].FileFormatType = FileFormatType.Png;
            // Optionally display the attachment as an icon
            sheet.OleObjects[oleIndex].DisplayAsIcon = true;

            // Create PDF save options and enable embedding of attachments
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            pdfOptions.EmbedAttachments = true;

            // Save the workbook as PDF with the embedded image attachment
            string outputPdf = "WorkbookWithEmbeddedImage.pdf";
            workbook.Save(outputPdf, pdfOptions);

            // Clean up the temporary image file
            File.Delete(imagePath);

            Console.WriteLine($"PDF saved to '{outputPdf}' with the image embedded as an attachment.");
        }
    }
}
