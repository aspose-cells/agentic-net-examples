// Title: Benchmark the impact of culture localization on chart PDF export performance for large workbooks using Aspose.Cells in C#
// Common Searches: how to measure export time of localized charts with Aspose.Cells in C# | benchmark chart PDF generation with different CultureInfo settings using Aspose.Cells | performance test exporting a large workbook to PDF after applying French culture to charts | timing Aspose.Cells PDF export when charts are localized for big spreadsheets | compare default and localized chart export speed in Aspose.Cells .NET
// Tags: chart localization performance Aspose.Cells | large workbook PDF export timing | cultureinfo effect on chart rendering | benchmark chart export .NET | measure Aspose.Cells export speed

using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a 5,000‑row by 20‑column workbook, adds a column chart, then exports the workbook to PDF twice—once with the invariant culture and once with French (fr‑FR) culture—recording and displaying the elapsed time for each export to evaluate the performance impact of chart localization.
class ChartLocalizationPerformance
{
    static void Main()
    {
        try
        {
            // Create a large workbook with sample data (5,000 rows × 20 columns)
            Workbook workbook = CreateLargeWorkbook(5000, 20);

            // Add a sample chart to the workbook
            AddSampleChart(workbook);

            // Export with the default (invariant) culture and measure duration
            TimeSpan defaultExportTime = ExportWorkbook(workbook, "default.pdf", CultureInfo.InvariantCulture);
            Console.WriteLine($"Export time with default culture: {defaultExportTime.TotalSeconds} seconds");

            // Export with a localized culture (French) and measure duration
            TimeSpan localizedExportTime = ExportWorkbook(workbook, "localized.pdf", new CultureInfo("fr-FR"));
            Console.WriteLine($"Export time with French culture: {localizedExportTime.TotalSeconds} seconds");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Generates a workbook filled with numeric data
    static Workbook CreateLargeWorkbook(int rowCount, int columnCount)
    {
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        for (int row = 0; row < rowCount; row++)
        {
            for (int col = 0; col < columnCount; col++)
            {
                ws.Cells[row, col].PutValue(row * columnCount + col + 1);
            }
        }

        return wb;
    }

    // Adds a simple column chart based on the first 10 rows of data
    static void AddSampleChart(Workbook wb)
    {
        Worksheet ws = wb.Worksheets[0];

        // Add a column chart positioned at cells A1 to J30
        int chartIndex = ws.Charts.Add(ChartType.Column, 0, 0, 30, 10);
        Chart chart = ws.Charts[chartIndex];

        // Define the series values (A1:A10) and categories (B1:B10)
        chart.NSeries.Add("A1:A10", true);
        chart.NSeries[0].Name = "Series 1";
        chart.NSeries.CategoryData = "B1:B10";
    }

    // Exports the workbook to PDF using the specified culture and returns the elapsed time
    static TimeSpan ExportWorkbook(Workbook wb, string outputPath, CultureInfo culture)
    {
        try
        {
            // Ensure the output directory exists
            string directory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Apply the culture to affect number/date formatting in charts
            wb.Settings.CultureInfo = culture;

            Stopwatch sw = Stopwatch.StartNew();

            // Export to PDF (could be any other format as needed)
            wb.Save(outputPath, SaveFormat.Pdf);

            sw.Stop();
            return sw.Elapsed;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to export workbook to '{outputPath}': {ex.Message}");
            return TimeSpan.Zero;
        }
    }
}
