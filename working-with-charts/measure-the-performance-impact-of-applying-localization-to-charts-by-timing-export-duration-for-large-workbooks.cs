// Title: Benchmark PDF export time for a 5,000‑row Aspose.Cells workbook before and after applying French chart localization (C#)
// AI Prompts: Generate C# code that creates a 5,000‑row worksheet, adds 20 column charts, saves the workbook as a PDF file, then sets Workbook.Settings.CultureInfo to "fr-FR", updates each chart title to a French string, saves again, and prints the elapsed milliseconds for both saves. | Write C# that repeats the same export timing test using the German locale ("de-DE") and outputs the workbook as PNG images instead of PDF, measuring the duration for each localized export. | Create a C# utility that runs the export benchmark for multiple locales (e.g., en-US, fr-FR, de-DE), varies the number of charts, records the timings in a CSV file, and optionally logs the results to the console.
// Common Searches: how to time Aspose.Cells PDF export with localized chart titles in C# | performance difference between localized and non‑localized chart export using Aspose.Cells | measure Aspose.Cells workbook export speed after setting CultureInfo | measure impact of French chart localization on PDF generation with Aspose.Cells | C# code to compare export times of large workbook with and without chart localization
// Tags: Aspose.Cells PDF creation latency | chart localization performance Aspose.Cells | large workbook export benchmark C# | Workbook.Settings.CultureInfo impact | measure chart title localization latency

using System;
using System.Diagnostics;
using System.Globalization;
using Aspose.Cells;
using Aspose.Cells.Charts;

// // Demonstrates creating a 5,000‑row workbook with 20 column charts, exporting to PDF once without localization and once after setting Workbook.Settings.CultureInfo to French and updating chart titles, measuring and printing the elapsed milliseconds for each export.
class ChartLocalizationPerformance
{
    static void Main()
    {
        try
        {
            // Create a large workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate worksheet with a sizable data set
            int totalRows = 5000;
            int totalCols = 5;
            for (int row = 0; row < totalRows; row++)
            {
                for (int col = 0; col < totalCols; col++)
                {
                    sheet.Cells[row, col].PutValue(row * col);
                }
            }

            // Add multiple charts to the worksheet
            int chartCount = 20;
            for (int i = 0; i < chartCount; i++)
            {
                int chartIndex = sheet.Charts.Add(ChartType.Column, 0, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];
                // Set the data range for the chart
                chart.NSeries.Add($"A1:E{totalRows}", true);
                // Set a default title
                chart.Title.Text = $"Sample Chart {i + 1}";
            }

            // Export without applying localization
            Stopwatch swNoLoc = new Stopwatch();
            swNoLoc.Start();
            workbook.Save("Export_NoLocalization.pdf", SaveFormat.Pdf);
            swNoLoc.Stop();
            Console.WriteLine($"Export without localization: {swNoLoc.ElapsedMilliseconds} ms");

            // Apply localization settings to the workbook and charts
            workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

            // Optionally localize chart titles
            foreach (Chart chart in sheet.Charts)
            {
                chart.Title.Text = "Graphique Exemple";
            }

            // Export with localization applied
            Stopwatch swLoc = new Stopwatch();
            swLoc.Start();
            workbook.Save("Export_WithLocalization.pdf", SaveFormat.Pdf);
            swLoc.Stop();
            Console.WriteLine($"Export with localization: {swLoc.ElapsedMilliseconds} ms");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
