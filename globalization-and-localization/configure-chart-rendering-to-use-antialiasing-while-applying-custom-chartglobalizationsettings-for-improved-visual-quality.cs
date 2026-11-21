// Title: How to enable chart anti‑aliasing with ChartGlobalizationSettings in Aspose.Cells for .NET (feature not available)
// Common Searches: asp.net enable anti-aliasing for Excel charts using Aspose.Cells | apply ChartGlobalizationSettings to improve chart rendering quality in C# | set anti-aliasing option for column chart with Aspose.Cells .NET | configure chart globalization settings for Excel files in Aspose.Cells C# | why Aspose.Cells chart anti-aliasing is unavailable in current version
// Tags: chart anti-aliasing Aspose.Cells .NET | ChartGlobalizationSettings usage C# | Aspose.Cells chart rendering quality | Excel chart anti-aliasing configuration | Aspose.Cells chart globalization API

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The sample creates a workbook, populates data, adds a column chart, and notes that ChartGlobalizationSettings for anti‑aliasing are not supported in the current Aspose.Cells version, then saves the workbook as ChartWithAntiAliasing.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);          // Values
            chart.NSeries.CategoryData = "A2:A4";     // Categories

            // Note: Chart globalization settings (anti‑aliasing, custom formats) are not
            // available in the current Aspose.Cells version, so this part is omitted.

            // Save the workbook
            string outputPath = "ChartWithAntiAliasing.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
