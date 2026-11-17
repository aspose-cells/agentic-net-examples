// Title: How to benchmark PDF export time for a 500‑sheet workbook using Aspose.Cells for .NET
// Common Searches: measure Aspose.Cells PDF save performance for a workbook with 500 worksheets in C# | C# benchmark exporting all worksheets to PDF with Aspose.Cells | how long does Aspose.Cells take to convert a 500‑sheet Excel file to PDF | performance test for Aspose.Cells PdfSaveOptions ExportAllWorksheets flag | timing large workbook PDF conversion using Aspose.Cells .NET API
// Tags: Aspose.Cells PDF export performance | benchmark large workbook conversion | PdfSaveOptions ExportAllWorksheets usage | C# timing workbook to PDF | 500 worksheet scalability test

using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample creates a workbook containing 500 worksheets, writes a simple value to each sheet, configures PdfSaveOptions (ExportAllWorksheets is true by default), measures the elapsed time with Stopwatch while saving the workbook to a PDF file, and outputs the duration and file path.
class WorkbookExportBenchmark
{
    static void Main()
    {
        try
        {
            // Create a new workbook (contains one default worksheet)
            Workbook workbook = new Workbook();

            // Add additional worksheets up to a total of 500
            for (int i = 1; i < 500; i++)
            {
                Worksheet sheet = workbook.Worksheets[workbook.Worksheets.Add()];
                // Populate each sheet with minimal data to avoid completely empty sheets
                sheet.Cells["A1"].PutValue($"Sheet {i + 1}");
            }

            // Configure PDF save options. By default all worksheets are exported.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Measure the time taken to export the workbook to PDF
            Stopwatch sw = Stopwatch.StartNew();

            // Save the workbook as PDF
            string outputPath = "ExportAllWorksheetsBenchmark.pdf";
            workbook.Save(outputPath, pdfOptions);

            sw.Stop();

            Console.WriteLine($"Exporting 500 worksheets took {sw.Elapsed.TotalSeconds:F2} seconds.");
            Console.WriteLine($"PDF saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
