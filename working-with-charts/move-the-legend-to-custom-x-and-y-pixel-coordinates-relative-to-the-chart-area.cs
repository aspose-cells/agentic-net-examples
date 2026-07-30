// Title: Move an Aspose.Cells chart legend to custom X/Y pixel coordinates in C#
// Description: Creates a workbook, adds sample data and a column chart, sets the legend to a free (NotDocked) position, and positions it using XPixel and YPixel offsets relative to the chart area, with optional overlay disabling.
// Keywords: Aspose.Cells chart legend position | C# legend XPixel YPixel | NotDocked legend Aspose.Cells | custom legend coordinates .NET | prevent legend overlay chart
// Common Searches: Aspose.Cells set legend pixel position | C# move chart legend to exact coordinates | NotDocked legend Aspose.Cells example | how to disable legend overlay in Aspose.Cells | custom legend placement in Excel using Aspose
// Developer Intent: Place a chart legend at precise pixel offsets within the chart area using Aspose.Cells for .NET.
// Use Cases: Avoid data series overlap by positioning the legend away from the plot area. | Maintain consistent legend placement across multiple generated charts. | Align the legend with other UI elements such as titles or annotations.
// AI Prompts: Generate C# code that sets an Aspose.Cells chart legend to (150, 60) pixels and disables overlay. | Show how to use Legend.Position = NotDocked, then assign XPixel and YPixel for a pie chart in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds sample data and a column chart, sets the legend to a free (NotDocked) position, and positions it using XPixel and YPixel offsets relative to the chart area, with optional overlay disabling.
class MoveLegendCustomPosition
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Set legend to a free (not docked) position so pixel coordinates are respected
        chart.Legend.Position = LegendPositionType.NotDocked;

        // Move legend to custom X and Y pixel coordinates relative to the chart area
        chart.Legend.XPixel = 120; // X offset in pixels
        chart.Legend.YPixel = 80;  // Y offset in pixels

        // Optional: prevent the legend from overlapping the chart
        chart.Legend.IsOverLay = false;

        // Save the workbook
        workbook.Save("LegendCustomPosition.xlsx");
    }
}
