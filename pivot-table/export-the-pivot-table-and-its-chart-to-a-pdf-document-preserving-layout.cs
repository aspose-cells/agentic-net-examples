// Title: Export an Excel PivotTable and its associated chart to PDF while preserving sheet layout using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export pivot table with chart to PDF preserving layout | How to keep original sheet formatting when saving an Excel pivot chart to PDF in .NET | Aspose.Cells PdfSaveOptions configuration for pivot tables and charts | Refresh all pivot tables before converting an Excel workbook to PDF using Aspose | Export Excel workbook containing a pivot chart to PDF without scaling in C#
// Tags: Aspose.Cells PDF layout preservation options | export pivot table and chart to PDF C# | refresh pivot tables Aspose.Cells before PDF conversion | OnePagePerSheet false Aspose.Cells setting | PdfCompliance PdfA1b for Excel to PDF export

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;          // Required for PivotTable class
using Aspose.Cells.Rendering;     // Required for PdfSaveOptions

// The example loads an Excel workbook, refreshes every PivotTable to reflect current data, configures PdfSaveOptions with OnePagePerSheet and AllColumnsInOnePagePerSheet set to false and PDF/A‑1b compliance, and saves the workbook—including the pivot table and its linked chart—to a PDF file while maintaining the original sheet layout.
class ExportPivotAndChartToPdf
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook that contains the pivot table and its associated chart.
            Workbook workbook = new Workbook(inputPath);

            // Refresh all pivot tables to ensure they reflect the latest data.
            foreach (Worksheet ws in workbook.Worksheets)
            {
                foreach (PivotTable pt in ws.PivotTables)
                {
                    pt.RefreshData();
                    pt.CalculateData();
                }
            }

            // Configure PDF save options to preserve the original layout of the sheet.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Keep the sheet layout as is; do not force everything onto a single page.
                OnePagePerSheet = false,
                AllColumnsInOnePagePerSheet = false,
                // Preserve the visual appearance of charts and pivot tables.
                Compliance = PdfCompliance.PdfA1b
            };

            // Save the workbook (including the pivot table and chart) to a PDF file.
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
