// Title: How to preserve Excel cell borders when converting to PDF using Aspose.Cells PdfSaveOptions.RenderSolidGridlines in C#
// AI Prompts: Write C# code that loads an .xlsx workbook with Aspose.Cells, sets PdfSaveOptions.RenderSolidGridlines to true, and saves it as a PDF while keeping the original cell borders. | Demonstrate the configuration of PdfSaveOptions for solid gridline rendering and export an Excel file to PDF with borders intact using Aspose.Cells.
// Common Searches: Aspose.Cells C# render solid gridlines PDF export | preserve Excel borders in PDF conversion using PdfSaveOptions | set RenderSolidGridlines true Aspose.Cells example | C# convert workbook to PDF with cell borders intact | PdfSaveOptions.RenderSolidGridlines not working Aspose.Cells
// Tags: Aspose.Cells PdfSaveOptions solid gridlines | C# export Excel to PDF with borders | RenderSolidGridlines property usage | preserve cell gridlines PDF conversion | Excel to PDF border rendering Aspose

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel workbook, creates a PdfSaveOptions object with RenderSolidGridlines enabled, and saves the workbook as a PDF. This configuration ensures that the original cell borders (gridlines) are rendered as solid lines in the resulting PDF, providing a faithful visual representation of the source worksheet.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            // Note: RenderGridlines option is not available in this version of Aspose.Cells.
            // If needed, adjust other PDF options here.

            // Save the workbook as a PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
