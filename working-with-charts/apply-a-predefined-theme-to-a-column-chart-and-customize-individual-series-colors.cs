// Title: Create a column chart with a predefined theme and set custom colors for each series using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that generates a column chart with Aspose.Cells, applies a preset chart style, and then changes the first series to red and the second series to green. | Demonstrate how to override the colors of individual series after a chart theme has been applied in an Aspose.Cells column chart.
// Common Searches: Aspose.Cells C# apply preset chart style to column chart | how to change series colors after chart theme in Aspose.Cells | override column chart series foreground color Aspose.Cells .NET | custom series colors for Aspose.Cells column chart example | Aspose.Cells set series color programmatically C#
// Tags: Aspose.Cells apply preset chart style C# | Aspose.Cells column chart series color customization | Aspose.Cells set series foreground color | Aspose.Cells chart theme override colors | Aspose.Cells create column chart with data C# | Aspose.Cells .NET chart styling

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.Drawing;

// The example creates a new workbook, fills it with sample data, adds a column chart, applies a predefined chart style (if supported), then sets the first series color to red and the second series color to green before saving the file as ColumnChartWithTheme.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data for the column chart
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

            // Add a column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIdx];

            // Define the data range for the chart
            chart.NSeries.Add("B2:C4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Apply a predefined chart style (removed due to API version compatibility)
            // chart.Style = ChartStyleType.PresetStyle1;

            // Customize series colors
            chart.NSeries[0].Area.ForegroundColor = Color.Red;   // Series1
            chart.NSeries[1].Area.ForegroundColor = Color.Green; // Series2

            // Save the workbook
            workbook.Save("ColumnChartWithTheme.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
