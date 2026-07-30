// Title: Aspose.Cells .NET: Parallel Export of Localized Charts to PDF with Low Memory Footprint
// Description: Shows how to create a workbook for each culture, apply MemorySetting.MemoryPreference, enable MultiThreadReading, fill month‑wise sales data via CultureInfo, build a column chart, and export the chart to a locale‑named PDF inside a Parallel.ForEach loop. Each workbook is disposed after export to keep memory usage stable in multi‑threaded environments.
// Keywords: Aspose.Cells | .NET | parallel chart export | localized charts | PDF export | MemorySetting.MemoryPreference | MultiThreadReading | culture-specific data | low memory processing | batch workbook generation
// Common Searches: Aspose.Cells export chart in parallel | memory efficient chart generation Aspose.Cells | localized PDF chart Aspose.Cells .NET | how to use MemorySetting.MemoryPreference | MultiThreadReading example Aspose.Cells | parallel processing of workbooks Aspose.Cells
// Developer Intent: Generate and export culture‑specific charts concurrently while minimizing memory consumption.
// Use Cases: Batch creation of monthly‑sales charts for multiple locales in a reporting service. | High‑throughput web API that returns localized PDF chart files without exhausting server RAM. | Automated international dashboard generation where each chart is produced in a separate thread. | Scheduled job that processes dozens of language versions of a chart while keeping the process lightweight.
// AI Prompts: Provide C# code to log Aspose.Cells memory usage before and after each chart export inside the parallel loop. | Explain how to share a single thread‑safe Random instance across Parallel.ForEach iterations to avoid duplicate values. | Show how to configure Aspose.Cells to write temporary files to a custom directory for better resource management. | Suggest ways to throttle the degree of parallelism to balance CPU usage and memory pressure.

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsMultiThreadedChartExport
{
    // Shows how to create a workbook for each culture, apply MemorySetting.MemoryPreference, enable MultiThreadReading, fill month‑wise sales data via CultureInfo, build a column chart, and export the chart to a locale‑named PDF inside a Parallel.ForEach loop. Each workbook is disposed after export to keep memory usage stable in multi‑threaded environments.
    public class LocalizedChartExporter
    {
        // List of locales (culture codes) to process – each will generate its own chart.
        private static readonly List<string> Locales = new List<string>
        {
            "en-US", "fr-FR", "de-DE", "es-ES", "ja-JP"
        };

        public static void Run()
        {
            // Process each locale in parallel while keeping memory usage stable.
            Parallel.ForEach(Locales, locale =>
            {
                // Create a new workbook for this locale.
                Workbook workbook = new Workbook();

                // Apply memory‑efficient settings.
                // MemoryPreference keeps the data model compact.
                workbook.Settings.MemorySetting = MemorySetting.MemoryPreference;

                // Enable multi‑thread reading on the cells collection.
                // This allows safe concurrent reads if needed later.
                workbook.Worksheets[0].Cells.MultiThreadReading = true;

                // Populate worksheet with localized sample data.
                PopulateWorksheet(workbook, locale);

                // Create a chart based on the data.
                Chart chart = CreateChart(workbook);

                // Export the chart to a PDF file named with the locale.
                string outputPath = Path.Combine("ExportedCharts", $"Chart_{locale}.pdf");
                Directory.CreateDirectory(Path.GetDirectoryName(outputPath));
                chart.ToPdf(outputPath);

                // Dispose the workbook to release any temporary files/resources.
                workbook.Dispose();

                Console.WriteLine($"Locale {locale}: chart exported to {outputPath}");
            });
        }

        private static void PopulateWorksheet(Workbook workbook, string locale)
        {
            // Example: create a simple table with localized month names.
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Header
            cells[0, 0].PutValue("Month");
            cells[0, 1].PutValue("Sales");

            // Sample data – 12 rows for months.
            var culture = new System.Globalization.CultureInfo(locale);
            for (int i = 1; i <= 12; i++)
            {
                // Use the culture to get month name.
                string monthName = culture.DateTimeFormat.GetMonthName(i);
                cells[i, 0].PutValue(monthName);
                // Random sales value.
                cells[i, 1].PutValue(new Random().Next(1000, 5000));
            }
        }

        private static Chart CreateChart(Workbook workbook)
        {
            Worksheet sheet = workbook.Worksheets[0];

            // Add a column chart covering the data range (A1:B13).
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart.
            chart.NSeries.Add("B2:B13", true);          // Values
            chart.NSeries.CategoryData = "A2:A13";     // Categories (month names)

            // Optional: give the chart a title.
            chart.Title.Text = "Monthly Sales";

            // Calculate the chart layout before exporting.
            chart.Calculate();

            return chart;
        }
    }

    // Entry point for demonstration.
    class Program
    {
        static void Main()
        {
            LocalizedChartExporter.Run();
        }
    }
}
