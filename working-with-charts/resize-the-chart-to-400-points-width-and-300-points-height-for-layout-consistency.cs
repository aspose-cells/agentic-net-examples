// Title: C# – Resize Aspose.Cells Chart to 400 pt × 300 pt
// Description: Shows how to create a workbook, add sample data, insert a column chart, and assign 400 points to ChartObject.Width and 300 points to ChartObject.Height before saving the file as ResizedChart.xlsx.
// Keywords: Aspose.Cells chart resize | ChartObject.Width C# | ChartObject.Height Aspose | set Excel chart dimensions programmatically | Aspose.Cells .NET chart size | adjust chart layout points | C# Excel chart width height | Aspose.Cells chart object size
// Common Searches: Aspose.Cells set chart width C# | How to change Excel chart height with Aspose.Cells | Specify chart size in points using Aspose.Cells .NET | C# code to adjust chart dimensions in a workbook | Programmatic chart layout control Aspose.Cells
// Developer Intent: Assign a width of 400 pt and a height of 300 pt to a chart object.
// Use Cases: Ensure uniform chart size across automated financial reports. | Fit charts into predefined page sections for printable Excel documents. | Create consistently sized graphics for embedding in PowerPoint or PDF exports.
// AI Prompts: Generate C# code that iterates through all charts in a worksheet and sets each to 400 pt width and 300 pt height using Aspose.Cells. | Explain the difference between ChartObject.Width/Height and the underlying Excel shape size properties. | Provide an example of resizing a chart after modifying its data series in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartResizeDemo
{
    // Shows how to create a workbook, add sample data, insert a column chart, and assign 400 points to ChartObject.Width and 300 points to ChartObject.Height before saving the file as ResizedChart.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
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

            // Add a column chart to the worksheet
            // Parameters: chart type, upper-left row, upper-left column, lower-right row, lower-right column
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Resize the chart: width = 400 points, height = 300 points
            chart.ChartObject.Width = 400;   // Width in points
            chart.ChartObject.Height = 300;  // Height in points

            // Save the workbook to a file
            workbook.Save("ResizedChart.xlsx");
        }
    }
}
