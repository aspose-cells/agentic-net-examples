// Title: C# – Parallel Localized Chart PDF Export with Compact Memory using Aspose.Cells
// Description: The sample creates a workbook, sets MemorySetting.MemoryPreference for a compact in‑memory model, enables MultiThreadReading, builds a single column chart, and launches a thread for each locale (en‑US, fr‑FR, de‑DE, ja‑JP, es‑ES). Each thread applies its CultureInfo, saves the shared chart as a locale‑named PDF, and signals a CountdownEvent before the workbook is disposed, keeping the memory footprint stable.
// Keywords: Aspose.Cells | C# chart export | parallel PDF generation | localized chart | MemoryPreference | MultiThreadReading | CountdownEvent | CultureInfo | low memory processing | thread‑safe chart rendering
// Common Searches: Aspose.Cells export chart to PDF multi thread | C# generate localized chart PDFs with Aspose.Cells | compact memory chart export Aspose.Cells .NET | parallel chart rendering Aspose.Cells example | set MemorySetting.MemoryPreference for chart export
// Developer Intent: Export a single chart to multiple locale‑specific PDF files concurrently while minimizing memory consumption.
// Use Cases: Produce sales charts in PDF for English, French, German, Japanese, and Spanish in a single batch to accelerate reporting. | Run high‑volume chart exports on a web server where each request needs its own culture formatting without blowing up RAM. | Process large workbooks with dozens of charts by reusing one chart object and leveraging Aspose.Cells low‑memory settings.
// AI Prompts: Write C# code that uses Aspose.Cells to export a chart to PDF for a list of locales in parallel, ensuring memory stays low with MemorySetting.MemoryPreference and MultiThreadReading. | Show how to switch the example to FileCache mode for extremely large workbooks while still supporting concurrent locale‑specific exports. | List best practices for disposing Aspose.Cells objects and handling exceptions during multi‑threaded chart PDF generation.

using System;
using System.Globalization;
using System.IO;
using System.Threading;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsMultiThreadedChartExport
{
    // The sample creates a workbook, sets MemorySetting.MemoryPreference for a compact in‑memory model, enables MultiThreadReading, builds a single column chart, and launches a thread for each locale (en‑US, fr‑FR, de‑DE, ja‑JP, es‑ES). Each thread applies its CultureInfo, saves the shared chart as a locale‑named PDF, and signals a CountdownEvent before the workbook is disposed, keeping the memory footprint stable.
    public class Program
    {
        // Entry point
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Use a memory‑efficient mode (compact in‑memory representation)
            // This reduces the overall memory footprint when many charts are processed.
            workbook.Settings.MemorySetting = MemorySetting.MemoryPreference;

            // Access the first worksheet and fill sample data
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Enable multi‑thread reading on the cells collection.
            // This allows concurrent read access to cell values while charts are being exported.
            cells.MultiThreadReading = true;

            // Populate data that will be used by all charts
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Value");
            for (int i = 2; i <= 6; i++)
            {
                cells[$"A{i}"].PutValue($"Item {i - 1}");
                cells[$"B{i}"].PutValue(i * 10);
            }

            // Add a single chart that will be reused for all locales.
            // The chart is created once to avoid repeated allocation of worksheet objects.
            int chartIndex = sheet.Charts.Add(ChartType.Column, 8, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B6", true);
            chart.NSeries.CategoryData = "A2:A6";
            chart.Title.Text = "Localized Sales Chart";

            // Define the locales for which the chart will be exported.
            string[] locales = new[] { "en-US", "fr-FR", "de-DE", "ja-JP", "es-ES" };

            // Use a countdown event to wait for all export threads to finish.
            CountdownEvent done = new CountdownEvent(locales.Length);

            foreach (string locale in locales)
            {
                Thread thread = new Thread(() =>
                {
                    try
                    {
                        // Set the current thread culture – this influences number/date formatting
                        // when the chart is rendered.
                        CultureInfo culture = new CultureInfo(locale);
                        Thread.CurrentThread.CurrentCulture = culture;
                        Thread.CurrentThread.CurrentUICulture = culture;

                        // Build a file name that reflects the locale.
                        string fileName = $"Chart_{locale}.pdf";

                        // Export the chart to PDF. The ToPdf(string) rule is used directly.
                        chart.ToPdf(fileName);

                        Console.WriteLine($"Exported chart for locale {locale} to {fileName}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error exporting chart for locale {locale}: {ex.Message}");
                    }
                    finally
                    {
                        // Signal that this thread has finished.
                        done.Signal();
                    }
                });

                // Start the export thread.
                thread.Start();
            }

            // Wait until all export threads have completed.
            done.Wait();

            // Dispose of the workbook to release any temporary files (important when using FileCache mode).
            workbook.Dispose();

            Console.WriteLine("All chart exports completed. Workbook resources released.");
        }
    }
}
