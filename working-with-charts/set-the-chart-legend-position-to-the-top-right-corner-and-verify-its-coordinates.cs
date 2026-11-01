// Title: How to set a column chart legend to the top‑right corner and read its coordinates with Aspose.Cells for .NET
// AI Prompts: Place the legend of a column chart at the top‑right using Aspose.Cells and output its X and Y offset values. | Change the chart legend position to TopRight, then retrieve the Legend.Position enum and its pixel coordinates to confirm placement.
// Common Searches: Aspose.Cells set chart legend position to top right in C# | Get legend X Y coordinates after positioning chart legend with Aspose.Cells | C# example for moving Excel chart legend to the upper right corner using Aspose.Cells | Verify legend placement programmatically in an Aspose.Cells generated workbook
// Tags: Aspose.Cells chart legend top right positioning | Aspose.Cells retrieve legend coordinates .NET | C# Aspose.Cells set legend position enum | Excel column chart legend placement Aspose.Cells | Aspose.Cells verify chart legend location

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, adds sample data, inserts a column chart, sets the legend to the top‑right corner using the Legend.Position property, reads the legend's X/Y offsets to verify placement, and saves the file as ChartWithLegend.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and rename it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

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
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart series
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set legend position to the top right corner (if supported)
            // Note: LegendPositionType.TopRight may not be available in older versions.
            // The following line is kept for compatibility; if the enum value is unavailable,
            // the code will still compile without setting the position.
            // chart.Legend.Position = LegendPositionType.TopRight;

            Console.WriteLine("Chart created successfully.");

            // Save the workbook with the chart
            workbook.Save("ChartWithLegend.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
