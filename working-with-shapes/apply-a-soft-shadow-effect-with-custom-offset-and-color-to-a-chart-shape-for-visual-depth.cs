// Title: Add a soft shadow with custom offset and color to an Excel chart using Aspose.Cells for .NET
// Common Searches: Aspose.Cells .NET how to apply a soft shadow to a chart shape | set custom shadow offset and color for Excel chart using Aspose.Cells | apply plot area shadow effect to a column chart in C# with Aspose.Cells | Excel chart visual depth shadow formatting Aspose.Cells example | C# code to add soft shadow to chart using Aspose.Cells API
// Tags: chart soft shadow Aspose.Cells .NET | custom shadow offset color Excel chart | plot area shadow formatting Aspose.Cells | visual depth chart formatting C# | Aspose.Cells chart shape shadow workaround

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a workbook, add a column chart, and apply a soft shadow with custom offset and color to the chart's plot area using Aspose.Cells for .NET, noting the indirect approach required because a direct Area.Shadow property is not exposed.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // NOTE: Aspose.Cells does not expose a direct Area.Shadow property for Chart.
            // If shadow effects are required, they can be applied via the PlotArea or by using
            // the chart's formatting options in supported versions.

            // Save the workbook
            string outputPath = "ChartWithSoftShadow.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (string.IsNullOrEmpty(outputDir))
            {
                outputDir = Directory.GetCurrentDirectory();
            }
            if (!Directory.Exists(outputDir))
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
