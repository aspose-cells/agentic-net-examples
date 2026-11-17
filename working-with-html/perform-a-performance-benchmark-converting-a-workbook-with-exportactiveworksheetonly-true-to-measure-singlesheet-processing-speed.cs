// Title: Benchmarking the export speed of the active worksheet to PDF with Aspose.Cells for .NET
// Common Searches: C# benchmark export of active worksheet to PDF using Aspose.Cells | measure performance of single-sheet PDF conversion with Aspose.Cells .NET | how to time Aspose.Cells ExportActiveWorksheetOnly PDF save | stopwatch example for active worksheet PDF export speed Aspose.Cells | performance test for large workbook single sheet PDF export in C#
// Tags: Aspose.Cells active worksheet PDF export benchmark | single sheet conversion speed Aspose.Cells | ExportActiveWorksheetOnly performance measurement | PdfSaveOptions timing test .NET | large workbook single sheet PDF conversion C#

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

namespace SingleSheetExportBenchmarkApp
{
    // The sample loads a large Excel workbook, isolates the active worksheet into a new workbook, and saves it as a PDF using Aspose.Cells' PdfSaveOptions. A Stopwatch records the elapsed time, providing a simple benchmark of single‑sheet export performance while handling missing files and ensuring the output directory exists.
    class SingleSheetExportBenchmark
    {
        static void Main()
        {
            // Paths for the source workbook and the exported PDF
            string inputPath = @"C:\Data\LargeWorkbook.xlsx";
            string outputPath = @"C:\Data\SingleSheetExport.pdf";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Start measuring time
                Stopwatch sw = Stopwatch.StartNew();

                // Load the source workbook
                Workbook sourceWb = new Workbook(inputPath);

                // Get the active worksheet
                int activeIndex = sourceWb.Worksheets.ActiveSheetIndex;
                Worksheet activeSheet = sourceWb.Worksheets[activeIndex];

                // Create a new workbook that will contain only the active worksheet
                Workbook singleSheetWb = new Workbook();

                // Remove the default empty sheet that comes with a new workbook
                if (singleSheetWb.Worksheets.Count > 0)
                {
                    singleSheetWb.Worksheets.RemoveAt(0);
                }

                // Add a copy of the active sheet to the new workbook
                // Use the sheet name overload to avoid version‑specific API issues
                singleSheetWb.Worksheets.AddCopy(activeSheet.Name);

                // Configure PDF save options (default options are sufficient)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Export the single‑sheet workbook to PDF
                singleSheetWb.Save(outputPath, pdfOptions);

                // Stop measuring time
                sw.Stop();

                Console.WriteLine($"Export of active worksheet completed in {sw.Elapsed.TotalSeconds:F3} seconds.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
