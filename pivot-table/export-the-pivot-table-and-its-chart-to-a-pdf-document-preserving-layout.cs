// Title: Export an Excel pivot table and its associated chart to PDF while preserving layout with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx workbook containing a pivot table and a linked chart, configures Aspose.Cells PdfSaveOptions to keep the original sheet layout, and saves the result as a PDF. | Demonstrate how to set PdfSaveOptions (OnePagePerSheet = false, AllColumnsInOnePagePerSheet = false) to export a workbook with a pivot chart without forcing a single‑page‑per‑sheet layout.
// Common Searches: Aspose.Cells C# export pivot table with chart to PDF preserving original layout | How to keep pivot chart formatting when converting Excel to PDF using Aspose.Cells | PdfSaveOptions settings to avoid one-page-per-sheet for pivot tables in Aspose.Cells | Convert Excel workbook containing pivot tables and charts to PDF in .NET without layout distortion
// Tags: export pivot table to PDF Aspose.Cells | preserve chart layout PdfSaveOptions | Aspose.Cells pivot chart PDF conversion | C# Excel to PDF with pivot table | PdfSaveOptions OnePagePerSheet false

using System;
using System.IO;
using Aspose.Cells;

// Loads an existing .xlsx workbook that includes a pivot table and its chart, configures PdfSaveOptions to retain the original sheet layout (disabling one‑page‑per‑sheet), and saves the workbook as a PDF using Aspose.Cells for .NET.
class ExportPivotAndChartToPdf
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input workbook exists to avoid FileNotFoundException.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the existing workbook that contains the pivot table and its chart.
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options to preserve the original layout.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = false,
                AllColumnsInOnePagePerSheet = false
                // Charts are rendered by default; no explicit RenderChart property needed.
            };

            // Save the workbook as a PDF using the configured options.
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully exported to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors (e.g., loading, saving) gracefully.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
