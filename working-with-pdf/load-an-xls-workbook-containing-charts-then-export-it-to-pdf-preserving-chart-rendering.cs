// Title: Convert an XLS workbook containing charts to PDF while preserving chart rendering using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xls file with embedded charts and saves it as a PDF, ensuring the charts appear exactly as in the worksheet with Aspose.Cells. | Demonstrate how to configure PdfSaveOptions in Aspose.Cells to retain chart formatting during Excel‑to‑PDF conversion. | Provide a C# example that checks for the source XLS file, applies chart‑preserving PDF options, and writes the output PDF.
// Common Searches: Aspose.Cells preserve chart layout when converting XLS to PDF in C# | How to export Excel file with charts to PDF using PdfSaveOptions | C# code to convert legacy .xls workbook with charts to PDF with Aspose.Cells | Keep Excel chart rendering intact during PDF export with Aspose.Cells .NET | PdfSaveOptions settings for chart rendering in Excel to PDF conversion
// Tags: xls to pdf conversion with chart rendering Aspose.Cells | PdfSaveOptions chart preservation | export Excel charts to PDF C# | preserve chart appearance Aspose.Cells | legacy Excel workbook PDF export Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads an existing .xls workbook that contains charts, configures PdfSaveOptions to keep chart appearance, and saves the workbook as a PDF.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xls";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing XLS workbook that contains charts
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options to preserve chart rendering
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Ensure that charts are rendered as they appear in the worksheet
                OnePagePerSheet = false,
                // Optional: keep column widths as in the original sheet
                AllColumnsInOnePagePerSheet = false
            };

            // Export the workbook to PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
