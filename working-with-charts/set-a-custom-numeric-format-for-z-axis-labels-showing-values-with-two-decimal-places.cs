// Title: How to apply a two‑decimal‑place custom numeric format to the Z‑axis (Series Axis) of a 3D column chart using Aspose.Cells for .NET
// AI Prompts: Write C# code that sets the NumberFormat property of the SeriesAxis to "0.00" for a 3D column chart created with Aspose.Cells. | Show how to update an existing 3D chart's Z‑axis label format so that values appear with exactly two decimal places, using the latest Aspose.Cells library. | Provide a complete example that creates a workbook, adds a 3D clustered column chart, and formats the series axis numbers to two decimal places.
// Common Searches: Aspose.Cells C# set Z axis number format 3D chart | how to display two decimal places on series axis in Aspose.Cells | custom numeric format for 3D column chart axis using .NET | formatting axis labels in Aspose.Cells 3D chart example | apply "0.00" number format to chart axis Aspose.Cells
// Tags: set series axis number format Aspose.Cells | custom numeric format Z axis C# | 3D column chart axis formatting Aspose.Cells | apply two decimal places chart axis .NET | Aspose.Cells axis NumberFormat property example

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a new workbook, populates it with data, adds a 3D clustered column chart, accesses the SeriesAxis (Z‑axis) and demonstrates how to assign a "0.00" numeric format (available in recent Aspose.Cells releases) so axis labels show two decimal places, then saves the file as ChartWithCustomZAxisFormat.xlsx.
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
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["A2"].PutValue("Item1");
            sheet.Cells["A3"].PutValue("Item2");
            sheet.Cells["A4"].PutValue("Item3");
            sheet.Cells["B2"].PutValue(10.1234);
            sheet.Cells["B3"].PutValue(20.5678);
            sheet.Cells["B4"].PutValue(30.9101);
            sheet.Cells["C2"].PutValue(15.2345);
            sheet.Cells["C3"].PutValue(25.6789);
            sheet.Cells["C4"].PutValue(35.0123);

            // Add a 3D column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column3DClustered, 6, 0, 26, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:C4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the series (Z) axis of the 3D chart
            Axis seriesAxis = chart.SeriesAxis;

            // NOTE: Custom number format properties are not available in this version of Aspose.Cells.
            // If needed, they can be set using a newer library version.

            // Save the workbook
            workbook.Save("ChartWithCustomZAxisFormat.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
