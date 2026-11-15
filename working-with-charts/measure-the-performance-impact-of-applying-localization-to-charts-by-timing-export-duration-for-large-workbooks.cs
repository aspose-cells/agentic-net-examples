// Title: Benchmark PDF export time for localized charts in a large Aspose.Cells workbook (C#)
// Common Searches: C# Aspose.Cells measure PDF export time for charts with different CultureInfo | how to benchmark chart localization impact on PDF generation using Aspose.Cells | performance test exporting a large workbook to PDF with French locale in Aspose.Cells | timing chart rendering when changing workbook.Settings.CultureInfo in Aspose.Cells .NET | compare default and localized chart export speed in Aspose.Cells C#
// Tags: chart PDF export performance Aspose.Cells | localization impact on chart rendering Aspose.Cells | large workbook export benchmark C# | CultureInfo setting performance Aspose.Cells | Aspose.Cells chart localization timing

using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a 20,000‑row by 10‑column workbook, adds a column chart, and saves the workbook to PDF twice—once with the default culture and once with French (fr-FR) culture—while measuring and outputting the elapsed milliseconds for each export to evaluate the performance effect of chart localization.
class ChartLocalizationPerformance
{
    static void Main()
    {
        try
        {
            // Parameters for a large workbook
            const int totalRows = 20000;   // number of data rows
            const int totalCols = 10;      // number of data columns

            // Create a new workbook and get the first worksheet
            using (Workbook workbook = new Workbook())
            {
                Worksheet sheet = workbook.Worksheets[0];

                // Fill the worksheet with sample data
                for (int row = 0; row < totalRows; row++)
                {
                    for (int col = 0; col < totalCols; col++)
                    {
                        sheet.Cells[row, col].PutValue(row + col);
                    }
                }

                // Add a column chart that uses the generated data
                // The chart occupies cells A1:J20 (adjust as needed)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 0, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart (first column as categories, remaining columns as series)
                string categoryRange = $"A1:A{totalRows}";
                string seriesRange = $"B1:J{totalRows}";
                chart.NSeries.Add(seriesRange, true);
                chart.NSeries.CategoryData = categoryRange;

                // Optional: set a title to see localization effect on text rendering
                chart.Title.Text = "Sample Data Chart";

                // Measure export duration without any localization (default culture)
                Stopwatch sw = new Stopwatch();
                sw.Start();
                workbook.Save("Export_Default.pdf", SaveFormat.Pdf);
                sw.Stop();
                Console.WriteLine($"Export time (default culture): {sw.ElapsedMilliseconds} ms");

                // Apply localization (e.g., French) and measure export duration
                workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

                sw.Restart();
                workbook.Save("Export_French.pdf", SaveFormat.Pdf);
                sw.Stop();
                Console.WriteLine($"Export time (French culture): {sw.ElapsedMilliseconds} ms");
            }
        }
        catch (Exception ex)
        {
            // Log the exception details for troubleshooting
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
