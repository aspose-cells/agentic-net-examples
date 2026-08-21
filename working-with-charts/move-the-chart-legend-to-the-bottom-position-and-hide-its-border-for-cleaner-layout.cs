// Title: Aspose.Cells .NET: Move Chart Legend to Bottom and Hide Its Border
// Description: Creates a workbook, adds a column chart, positions the legend at the bottom, disables the legend border, and saves the file as ChartWithBottomLegend.xlsx using Aspose.Cells for C#.
// Keywords: Aspose.Cells | .NET chart legend position | legend bottom Aspose.Cells | hide legend border C# | Excel chart formatting Aspose | C# Aspose.Cells legend customization
// Common Searches: Aspose.Cells set legend to bottom .NET | remove legend border Aspose.Cells C# | chart legend position bottom Aspose.Cells | hide chart legend border Aspose.Cells example | C# Aspose.Cells chart legend formatting
// Developer Intent: Place the chart legend at the bottom and hide its border.
// Use Cases: Design clean Excel reports where legends are positioned below the chart for better readability. | Apply corporate styling that requires legend borders to be invisible across multiple charts. | Automate generation of dashboards with consistently formatted legends in a .NET application.
// AI Prompts: Show C# code to set a chart legend to the bottom and hide its border using Aspose.Cells. | Give an example of customizing legend position and border visibility for any chart type in Aspose.Cells .NET. | Explain how to programmatically adjust legend properties for multiple charts in a workbook with Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds a column chart, positions the legend at the bottom, disables the legend border, and saves the file as ChartWithBottomLegend.xlsx using Aspose.Cells for C#.
class Program
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
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Move the legend to the bottom of the chart
        chart.Legend.Position = LegendPositionType.Bottom;

        // Hide the legend border for a cleaner layout
        chart.Legend.Border.IsVisible = false;

        // Save the workbook with the modified chart
        workbook.Save("ChartWithBottomLegend.xlsx");
    }
}
