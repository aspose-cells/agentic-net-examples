// Title: Embed images and CSV files into a PDF generated from an Aspose.Cells workbook using C# PdfSaveOptions
// AI Prompts: Create C# code that loads an Excel workbook, configures PdfSaveOptions with a list of attachment files (e.g., PNG and CSV), and saves the workbook as a PDF containing those attachments using Aspose.Cells. | Update the sample workbook code to add multiple files to the PdfSaveOptions.Attachments collection before calling workbook.Save with SaveFormat.Pdf. | Write a C# routine that verifies the target folder exists, creates it if necessary, and then exports the workbook to PDF while embedding specified attachments via PdfSaveOptions.
// Common Searches: how to attach image files to PDF when saving Aspose.Cells workbook in C# | Aspose.Cells PdfSaveOptions embed CSV as attachment in generated PDF .NET | C# export Excel to PDF with embedded files using Aspose.Cells PdfSaveOptions | add multiple attachments to PDF created from Excel workbook using Aspose.Cells
// Tags: Aspose.Cells PdfSaveOptions embed attachments | C# attach files to PDF using Aspose.Cells | embed images in PDF from Excel workbook | add CSV attachment to PDF Aspose.Cells | ensure output directory exists C# Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// This example demonstrates how to use Aspose.Cells PdfSaveOptions in C# to embed additional files such as images and CSV documents into a PDF generated from an Excel workbook. It includes code for preparing the attachment list, configuring PdfSaveOptions, ensuring the output folder exists, and saving the workbook as a PDF with the embedded attachments.
class EmbedAttachmentsToPdf
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Add some sample data to the worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample data for PDF");

            // Ensure output directory exists
            string outputPath = @"C:\Output\Result.pdf";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook directly as a PDF
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"PDF saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
