// Title: Parallel PDF Export of Localized Charts with Low Memory Using Aspose.Cells for .NET
// Description: Creates a separate workbook for each locale, applies MemorySetting.FileCache, enables MultiThreadReading, builds a column chart, and exports it to PDF inside a Parallel.ForEach loop. Each workbook is disposed immediately to keep the memory footprint stable during concurrent processing.
// Keywords: Aspose.Cells | C# | parallel chart export | PDF generation | MemorySetting.FileCache | MultiThreadReading | localized charts | low‑memory processing | thread‑safe Aspose.Cells | chart to PDF .NET
// Common Searches: Aspose.Cells export chart to PDF in parallel | reduce memory usage when exporting charts Aspose.Cells | enable MultiThreadReading for chart export C# | Parallel.ForEach Aspose.Cells PDF chart | generate localized chart PDFs with Aspose.Cells
// Developer Intent: Produce PDF files of locale‑specific charts concurrently while preventing high memory consumption.
// Use Cases: Batch creation of language‑specific sales charts for a multinational report. | Background service that renders thousands of chart PDFs without exhausting server RAM. | On‑demand generation of regional performance dashboards in a web application.
// AI Prompts: Show how to limit Parallel.ForEach degree of parallelism to protect the thread pool. | Add code that logs workbook memory usage before and after each export using Aspose.Cells diagnostics. | Demonstrate reusing a single workbook template for all locales while still applying MemorySetting.FileCache per export.

using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a separate workbook for each locale, applies MemorySetting.FileCache, enables MultiThreadReading, builds a column chart, and exports it to PDF inside a Parallel.ForEach loop. Each workbook is disposed immediately to keep the memory footprint stable during concurrent processing.
    public static class ChartExportProcessor
    {
        /// <summary>
        /// Exports a simple column chart for each locale to a PDF file using multiple threads.
        /// </summary>
        /// <param name="locales">Array of locale identifiers (e.g., "en-US", "fr-FR").</param>
        /// <param name="outputFolder">Folder where the PDF files will be saved.</param>
        public static void ExportChartsParallel(string[] locales, string outputFolder)
        {
            // Ensure the output directory exists.
            Directory.CreateDirectory(outputFolder);

            // Use Parallel.ForEach to run the export for each locale on a separate thread.
            Parallel.ForEach(locales, locale =>
            {
                // Create a new workbook (rule: create).
                Workbook workbook = new Workbook();

                // Set memory usage to FileCache to keep memory footprint low.
                // This setting will be the default for newly created worksheets.
                workbook.Settings.MemorySetting = MemorySetting.FileCache;

                // Access the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Enable multi‑thread reading on the cells collection.
                cells.MultiThreadReading = true;

                // Populate some sample data that could be localized.
                // In a real scenario you would load locale‑specific data here.
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Value");
                cells["A2"].PutValue($"{locale} - Item 1");
                cells["B2"].PutValue(10);
                cells["A3"].PutValue($"{locale} - Item 2");
                cells["B3"].PutValue(20);
                cells["A4"].PutValue($"{locale} - Item 3");
                cells["B4"].PutValue(30);

                // Add a column chart.
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart.
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Optional: calculate the chart layout before exporting.
                chart.Calculate();

                // Build the output file name.
                string pdfPath = Path.Combine(outputFolder, $"Chart_{locale}.pdf");

                // Export the chart to PDF (rule: ToPdf(string)).
                chart.ToPdf(pdfPath);

                // Release unmanaged resources promptly.
                workbook.Dispose();

                // Log completion (optional).
                Console.WriteLine($"Thread {Thread.CurrentThread.ManagedThreadId}: Exported chart for locale '{locale}' to '{pdfPath}'.");
            });
        }

        // Example entry point.
        public static void Main()
        {
            // Example list of locales to process.
            string[] locales = new[] { "en-US", "fr-FR", "de-DE", "es-ES", "ja-JP" };

            // Destination folder for the exported PDFs.
            string outputFolder = Path.Combine(Environment.CurrentDirectory, "ChartExports");

            // Run the parallel export.
            ExportChartsParallel(locales, outputFolder);

            Console.WriteLine("All chart exports completed.");
        }
    }
}
