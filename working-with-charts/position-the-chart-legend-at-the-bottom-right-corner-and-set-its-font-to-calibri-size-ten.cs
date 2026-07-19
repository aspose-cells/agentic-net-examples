// Title: Aspose.Cells for .NET – Move Chart Legend to Bottom‑Right and Apply Calibri 10 Font
// Description: Creates a workbook, adds a column chart, makes the legend visible, positions it at the lower‑right corner outside the plot area, and sets the legend font to Calibri size 10 before saving the file.
// Keywords: Aspose.Cells | .NET | C# | chart legend positioning | bottom right legend | Calibri font | font size 10 | LegendPositionType.Corner | disable legend overlay | Excel chart formatting
// Common Searches: Aspose.Cells place legend bottom right | set legend font Calibri Aspose.Cells | how to disable legend overlay in Aspose.Cells chart | LegendPositionType.Corner example C# | format chart legend Aspose.Cells .NET
// Developer Intent: Programmatically locate the chart legend at the lower‑right edge and change its typeface to Calibri 10.
// Use Cases: Design Excel dashboards where the legend must stay out of the data area. | Enforce a consistent visual style for charts across multiple reports. | Build reusable chart templates that automatically apply the preferred legend layout and typography.
// AI Prompts: Generate C# code using Aspose.Cells that adds a column chart, moves its legend to the lower‑right corner, turns off overlay, and sets the font to Calibri 10. | Explain how to modify an existing Aspose.Cells chart to change the legend position and font without recreating the chart. | Provide step‑by‑step instructions for using LegendPositionType.Corner and Font properties to style a chart legend in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds a column chart, makes the legend visible, positions it at the lower‑right corner outside the plot area, and sets the legend font to Calibri size 10 before saving the file.
class ChartLegendExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
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

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Ensure the legend is visible
        chart.ShowLegend = true;

        // Position the legend at the bottom‑right corner
        // Corner places the legend in the chart's corner; combined with Bottom it appears at the lower right.
        chart.Legend.Position = LegendPositionType.Corner;
        // Optionally, fine‑tune the position if needed
        chart.Legend.IsOverLay = false; // keep legend outside the plot area

        // Set legend font to Calibri, size 10
        chart.Legend.Font.Name = "Calibri";
        chart.Legend.Font.Size = 10;

        // Save the workbook
        workbook.Save("ChartWithBottomRightLegend.xlsx");
    }
}
