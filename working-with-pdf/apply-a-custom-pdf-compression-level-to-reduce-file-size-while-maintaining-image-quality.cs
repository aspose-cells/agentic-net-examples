// Title: Apply a custom PDF compression level and image quality settings when converting Excel to PDF with Aspose.Cells for .NET
// AI Prompts: Generate C# code that sets PdfSaveOptions.Compression to a specific PdfCompressionLevel and configures ImageCompression to preserve image quality while reducing PDF size. | Show how to adjust Aspose.Cells PDF export options to use high‑quality image compression (e.g., PdfImageCompression.Jpeg) together with a chosen compression level for smaller PDFs. | Provide a snippet that converts an .xlsx file to PDF with PdfSaveOptions configured for balanced file size and image fidelity.
// Common Searches: Aspose.Cells C# set PDF compression level to reduce file size | How to control image quality when saving Excel as PDF with Aspose.Cells | PdfSaveOptions Compression and ImageCompression properties example | Convert Excel to PDF with custom compression using Aspose.Cells .NET | Reduce PDF size from Excel workbook without losing images Aspose.Cells
// Tags: Aspose.Cells PdfSaveOptions compression level | Excel to PDF image compression Aspose.Cells | custom PDF size reduction .NET Aspose.Cells | set PdfCompressionLevel in Aspose.Cells | preserve image fidelity PDF export Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering; // Required for PdfSaveOptions

// The example loads an Excel workbook, creates a PdfSaveOptions object, sets the Compression property (e.g., PdfCompressionLevel.Normal) and ImageCompression (e.g., PdfImageCompression.Jpeg) to balance file size and image quality, then saves the workbook as a PDF while handling missing file and exception scenarios.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the source Excel file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the source Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (default settings)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
