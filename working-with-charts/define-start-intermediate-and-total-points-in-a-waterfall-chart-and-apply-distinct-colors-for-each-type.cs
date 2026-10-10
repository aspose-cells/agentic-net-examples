// Title: Set custom colors for start, intermediate, and total points in a waterfall chart with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to create a waterfall chart and assign a unique foreground color to the start point. | Show how to change the color of an intermediate point in an Aspose.Cells waterfall chart series. | Provide a complete example that colors the total point green in a waterfall chart generated with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# how to color individual points in a waterfall chart | set start point blue in Aspose.Cells waterfall chart example | change intermediate point color in Aspose.Cells waterfall series .NET | assign green color to total point of waterfall chart using Aspose.Cells | custom point colors waterfall chart Aspose.Cells tutorial
// Tags: set waterfall chart point foreground color Aspose.Cells | customize series point color .NET | apply distinct colors to start intermediate total points Aspose.Cells | waterfall chart point formatting C# | Aspose.Cells chart series point styling

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

// Demonstrates creating a workbook, adding data, inserting a waterfall chart, and using the Series.Points[i].Area.ForegroundColor property to color the start point blue, an intermediate point orange, and the total point green, then saving the file as WaterfallChart.xlsx.
class WaterfallChartExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data for the waterfall chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Start");
            sheet.Cells["B2"].PutValue(0);
            sheet.Cells["A3"].PutValue("Sales");
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["A4"].PutValue("Expenses");
            sheet.Cells["B4"].PutValue(-150);
            sheet.Cells["A5"].PutValue("Profit");
            sheet.Cells["B5"].PutValue(0);

            // Add a waterfall chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the series data range and categories
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Access the series (only one series in this case)
            Series series = chart.NSeries[0];

            // Apply distinct colors for each point using the Area.ForegroundColor property
            series.Points[0].Area.ForegroundColor = Color.Blue;     // Start point color
            series.Points[2].Area.ForegroundColor = Color.Orange;   // Intermediate point color
            series.Points[3].Area.ForegroundColor = Color.Green;    // Total point color

            // Save the workbook with the chart
            workbook.Save("WaterfallChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
