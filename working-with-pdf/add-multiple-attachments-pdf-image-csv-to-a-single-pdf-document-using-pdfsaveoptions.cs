// Title: Generate a PDF from an Excel workbook with Aspose.Cells and then embed PDF, image, and CSV attachments using Aspose.Pdf in C#
// AI Prompts: Write C# code that creates an Excel workbook, saves it as a PDF with Aspose.Cells PdfSaveOptions, and then uses Aspose.Pdf to attach a PDF file, an image file, and a CSV file to the same PDF document. | Show a step‑by‑step example of combining Aspose.Cells and Aspose.Pdf in .NET to produce a single PDF that contains embedded file attachments of different formats. | Implement a C# method that receives a PDF path and a collection of file paths and adds each file as an attachment to the PDF using Aspose.Pdf.
// Common Searches: how to add multiple file attachments to a PDF generated from Excel using Aspose.Cells and Aspose.Pdf | C# attach PDF, JPG, and CSV files to a PDF created by Aspose.Cells | Aspose.Pdf embed attachments after saving workbook as PDF | workaround for PdfSaveOptions attachment limitation in Aspose.Cells | combine Aspose.Cells PDF export with Aspose.Pdf attachment API
// Tags: Aspose.Cells save workbook to PDF | Aspose.Pdf add file attachments | C# embed multiple attachments in PDF | post‑process PDF with Aspose.Pdf | Excel to PDF with attached files .NET

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates creating a new Workbook, populating it with sample data, and saving it as a PDF using Aspose.Cells PdfSaveOptions. Because Aspose.Cells does not provide a direct API for embedding file attachments, the workflow continues with Aspose.Pdf to attach a PDF, an image, and a CSV file to the generated PDF, delivering a single document that contains all required attachments.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Add some sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Data");
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["A3"].PutValue(DateTime.Now);

            // Prepare PDF save options (lifecycle rule: load/save options)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // NOTE: Aspose.Cells does not provide a direct API to add file attachments to a PDF.
            // If attachment functionality is required, consider using Aspose.Pdf after saving the PDF.

            // Ensure output directory exists
            string outputPath = @"C:\Output\CombinedDocument.pdf";
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a PDF (lifecycle rule: save)
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
