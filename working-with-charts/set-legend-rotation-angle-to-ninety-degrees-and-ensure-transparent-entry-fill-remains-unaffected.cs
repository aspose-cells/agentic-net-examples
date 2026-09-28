// Title: Rotate a chart legend 90° in Aspose.Cells for .NET without changing its transparent entry fill
// AI Prompts: Create C# code with Aspose.Cells that sets Chart.Legend.RotationAngle to 90 degrees while leaving the legend entry fill transparency untouched. | Generate a .NET example that adds a column chart, rotates its legend by ninety degrees, and ensures the legend entries remain transparent.
// Common Searches: Aspose.Cells C# rotate chart legend 90 degrees | keep legend entry fill transparent when rotating legend Aspose.Cells | set legend rotation angle without affecting fill transparency .NET | example of chart legend rotation in Aspose.Cells for .NET
// Tags: chart legend rotation Aspose.Cells | preserve legend entry fill transparency .NET | set legend rotation angle C# | column chart legend customization Aspose.Cells | transparent legend fill Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds a column chart, sets the legend's RotationAngle to 90 degrees, and saves the file, keeping the legend entry fill transparent.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule)
            var workbook = new Workbook();

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Add a sample chart (using Column chart type)
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            var chart = sheet.Charts[chartIdx];

            // Sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);

            // Set data series and categories
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Set legend rotation angle to ninety degrees
            chart.Legend.RotationAngle = 90;

            // Save the workbook (lifecycle rule)
            string outputPath = "LegendRotation.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
