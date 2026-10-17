// Title: Convert an Excel workbook to PDF with cell background colors preserved using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export Excel to PDF keeping cell fill colors | How to retain background shading when saving .xlsx as PDF with Aspose.Cells | C# PdfSaveOptions to preserve cell background in PDF conversion | Generate PDF/A-1b from Excel while maintaining cell colors using Aspose.Cells | Save workbook to PDF without losing cell background colors in .NET
// Tags: Aspose.Cells PdfSaveOptions background colors | C# Excel to PDF preserving cell shading | PDF/A-1b compliance Aspose.Cells export | AllColumnsInOnePagePerSheet false PDF output | Workbook.Save PDF cell fill preservation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an existing XLSX file with Aspose.Cells, configures PdfSaveOptions (disabling AllColumnsInOnePagePerSheet and enabling PDF/A‑1b compliance), and saves the workbook as a PDF, ensuring that cell background colors are retained in the generated document.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Keep the original column widths and layout
                AllColumnsInOnePagePerSheet = false,
                // Set PDF/A compliance for better compatibility
                Compliance = PdfCompliance.PdfA1b
                // RenderGridLines property is not available for PdfSaveOptions; grid lines are rendered by default
            };

            // Export the workbook to PDF while preserving cell background colors
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
