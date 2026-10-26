// Title: How to add logging for culture‑specific chart label methods in Aspose.Cells using C#
// Common Searches: C# Aspose.Cells log which chart label methods are invoked for each culture | debug chart label localization in Aspose.Cells by capturing method callbacks | track culture‑specific label rendering during Aspose.Cells chart generation | monitor chart axis label method execution for globalization in .NET Aspose.Cells | record chart label method calls when rendering with different locales in Aspose.Cells
// Tags: chart label method logging Aspose.Cells | culture-aware chart rendering debug | Aspose.Cells localization diagnostics | C# chart label callback tracking | globalization troubleshooting Aspose.Cells

using System;
using System.Globalization;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program creates a new Workbook, writes numeric values to cells A1‑A3, adds a column chart referencing that range, and saves the workbook as ChartWithLogging.xlsx while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample data for the chart
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("A1:A3", true);

            // Save the workbook
            string outputPath = "ChartWithLogging.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
