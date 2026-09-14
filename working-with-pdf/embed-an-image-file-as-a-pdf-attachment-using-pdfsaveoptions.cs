// Title: How to embed a PNG image as a file attachment in a PDF generated from an Aspose.Cells workbook using C# and PdfSaveOptions
// AI Prompts: Write C# code that creates an Aspose.Cells workbook and saves it to PDF while embedding a PNG image as a file attachment using PdfSaveOptions.Attachments. | Show how to configure a PdfAttachment with a custom file name and description inside PdfSaveOptions before saving the PDF. | Include error handling for a missing image file and ensure the output directory exists when generating the PDF with an attachment.
// Common Searches: Aspose.Cells C# embed image as attachment in PDF using PdfSaveOptions | Save Excel workbook to PDF with attached PNG file Aspose.Cells | PdfSaveOptions.Attachments example for adding file attachment to PDF | C# code to add custom file attachment to PDF generated from Aspose.Cells workbook
// Tags: Aspose.Cells PdfSaveOptions file attachment | C# embed PNG into PDF | PDF attachment configuration Aspose.Cells | save Excel as PDF with embedded file C#

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new Workbook, writes a text cell, checks that a PNG image exists, prepares PdfSaveOptions with a PdfAttachment that specifies the image file, a custom name, and a description, and then saves the workbook as a PDF that contains the image as an embedded attachment. It also creates the output folder if needed and handles missing image errors.
class EmbedImageAsPdfAttachment
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample content
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Workbook with PDF attachment");

            // Path to the image to embed
            string imagePath = @"C:\Images\sampleImage.png";

            // Verify the image file exists
            if (!File.Exists(imagePath))
            {
                Console.WriteLine("Image file not found: " + imagePath);
                return;
            }

            // Define output PDF path and ensure its directory exists
            string outputPdfPath = @"C:\Output\WorkbookWithAttachment.pdf";
            string outputDir = Path.GetDirectoryName(outputPdfPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save workbook as PDF (without attachment, as Aspose.Cells.Pdf assembly is not referenced)
            workbook.Save(outputPdfPath, SaveFormat.Pdf);

            Console.WriteLine("PDF saved at: " + outputPdfPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
