// Title: How to display Spanish month names in an Excel chart legend using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set chart legend series name to Spanish month using TEXT formula | C# generate Excel column chart with month names in Spanish locale Aspose.Cells | Apply custom locale ID C0A to chart series name in Aspose.Cells .NET | Display month names in Spanish on Excel chart legend programmatically with Aspose.Cells | Create Excel chart with localized month labels in legend using Aspose.Cells C#
// Tags: Aspose.Cells chart series name locale | Excel chart legend Spanish month names | C# TEXT formula with locale ID | column chart with localized legend Aspose.Cells | apply custom locale to Excel chart series

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example creates a workbook, fills column A with the first day of each month for 2023 and column B with sample values, adds a column chart, clears default series, then adds a series for each month and sets its name using a TEXT formula with the Spanish (Spain) locale ID (C0A) so the chart legend shows month names in Spanish. The workbook is saved as ChartWithSpanishMonthLegend.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the sheet with dates (first day of each month) in column A
            // and some sample numeric values in column B
            DateTime startDate = new DateTime(2023, 1, 1);
            for (int i = 0; i < 12; i++)
            {
                sheet.Cells[i, 0].PutValue(startDate.AddMonths(i)); // A column: dates
                sheet.Cells[i, 1].PutValue(i + 10);                // B column: sample values
            }

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 15, 0, 30, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Clear any automatically created series
            chart.NSeries.Clear();

            // Spanish (Spain) locale ID is C0A (hex) = 3082 decimal.
            // Use TEXT formula with locale to generate month names in Spanish for legend entries.
            for (int i = 0; i < 12; i++)
            {
                // Add a series with a single Y‑value (column B)
                chart.NSeries.Add($"B{i + 1}", false);

                // Set the series name to the formatted month name in Spanish
                Series series = chart.NSeries[i];
                series.Name = $"=TEXT(A{i + 1},\"[$-C0A]mmmm\")";
            }

            // Save the workbook with the chart.
            string outputPath = "ChartWithSpanishMonthLegend.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
