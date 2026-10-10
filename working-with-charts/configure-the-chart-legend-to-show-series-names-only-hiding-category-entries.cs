// Title: Configure an Aspose.Cells column chart legend to display only series names and hide category entries in C#
// AI Prompts: Generate C# code using Aspose.Cells that creates a column chart and configures the legend to show only the series names, removing category labels. | Modify an existing Aspose.Cells chart in C# so that its legend entries are limited to series names while keeping the legend position unchanged.
// Common Searches: how to hide category labels in chart legend with Aspose.Cells C# | Aspose.Cells legend show only series names column chart | C# set Excel chart legend to series only using Aspose.Cells | remove category entries from Aspose.Cells chart legend programmatically | customize Aspose.Cells chart legend to display series names only
// Tags: Aspose.Cells column chart legend series-only | C# hide chart legend categories Aspose.Cells | Aspose.Cells set legend entries to series names | Excel chart legend customization Aspose.Cells .NET | Aspose.Cells chart legend configuration C#

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example creates a workbook, fills it with sample data, adds a column chart, and configures the chart legend to display only the series names while omitting the category (X‑axis) entries, then saves the file as ChartWithLegend.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Add a column chart positioned from row 5, column 0 to row 20, column 10
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set series values and category (X‑axis) data
            chart.NSeries.Add("B2:C4", true);          // Series values range
            chart.NSeries.CategoryData = "A2:A4";     // Category (X‑axis) range

            // Configure legend (default shows series names)
            chart.Legend.Position = LegendPositionType.Right; // Place legend on the right

            // Save the workbook with the chart
            string outputPath = "ChartWithLegend.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
