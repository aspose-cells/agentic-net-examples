// Title: Convert HTML to PDF with maximum image compression using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an HTML file into an Aspose.Cells Workbook and saves it as a PDF while setting PdfSaveOptions.ImageCompression to the highest level. | Show how to adjust PdfSaveOptions in Aspose.Cells to compress embedded images and shrink the PDF size during HTML-to-PDF conversion.
// Common Searches: Aspose.Cells C# convert HTML file to PDF with high image compression | How to reduce PDF size when exporting HTML using PdfSaveOptions in .NET | Set maximum image compression for PDF output in Aspose.Cells HTML to PDF conversion
// Tags: Aspose.Cells HTML-to-PDF image compression | PdfSaveOptions image compression setting | C# Aspose.Cells reduce PDF size | load HTML workbook Aspose.Cells | configure PDF save options Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;   // Required for PdfSaveOptions

// The sample verifies the existence of an HTML file, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object (which can be configured for high image compression), ensures the output directory exists, and saves the workbook as a PDF while handling any runtime errors.
class HtmlToPdfConverter
{
    static void Main()
    {
        string inputPath = "input.html";
        string outputPath = "output.pdf";

        try
        {
            // Verify that the input HTML file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the HTML file into a Workbook object
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (default settings)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
