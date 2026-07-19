// Title: Refresh Aspose.Cells Chart After Changing Data Label and Series Colors (C#)
// Description: Demonstrates how to modify data label font color and series color in a column chart, then recalculate the chart with a ChartExtensions.Refresh method (which calls Chart.Calculate) so the changes appear in the saved Excel file.
// Keywords: Aspose.Cells | Chart.Refresh extension | Chart.Calculate | C# chart styling | update chart colors | data label font color | column chart | Excel workbook | extension method | recalculate chart
// Common Searches: Aspose.Cells refresh chart after style change | C# update data label color in Excel chart | How to recalculate Aspose.Cells chart | ChartExtensions.Refresh Aspose.Cells example | Apply new series color and redraw chart
// Developer Intent: Recalculate a chart to display updated label fonts or series colors before saving the workbook.
// Use Cases: After programmatically changing data label font properties, call chart.Refresh to apply the visual updates. | When altering a series' foreground color, invoke chart.Refresh to ensure the new color is rendered in the chart image. | Wrap Chart.Calculate in a reusable ChartExtensions.Refresh method for consistent chart updates across projects.
// AI Prompts: Generate C# code that changes a chart's data label font size and refreshes the chart using Aspose.Cells. | Show how to implement a ChartExtensions.Refresh method that calls Calculate and use it after modifying series colors. | Explain how to guarantee that style changes to labels and series are saved in the Excel file with Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartRefreshDemo
{
    // Extension method to simulate Chart.Refresh by invoking Calculate()
    // Demonstrates how to modify data label font color and series color in a column chart, then recalculate the chart with a ChartExtensions.Refresh method (which calls Chart.Calculate) so the changes appear in the saved Excel file.
    public static class ChartExtensions
    {
        public static void Refresh(this Chart chart)
        {
            // Recalculates the chart layout and updates visual elements
            chart.Calculate();
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");

            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the first series
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;

            // Change the font color of the data labels
            series.DataLabels.Font.Color = Color.Green;
            // Apply the font settings to all data label instances
            series.DataLabels.ApplyFont();

            // Change the series color (example: set to a solid red)
            series.Area.ForegroundColor = Color.Red;

            // Refresh the chart display to reflect the modifications
            chart.Refresh();

            // Save the workbook
            workbook.Save("ChartRefreshDemo.xlsx");
        }
    }
}
