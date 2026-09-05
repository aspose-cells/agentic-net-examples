// Title: Resize a column chart to 400 × 300 points using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a workbook, adds sample data, inserts a column chart, and sets its ChartObject.Width to 400 and Height to 300 points with Aspose.Cells. | Show how to programmatically adjust the size of an Aspose.Cells chart object in points after it has been added to a worksheet. | Provide a complete Aspose.Cells example that resizes an existing chart to specific dimensions for consistent layout.
// Common Searches: Aspose.Cells C# set chart width to 400 points and height to 300 points | How to change chart size in an Excel file using Aspose.Cells .NET | Resize column chart programmatically with Aspose.Cells for .NET | Set ChartObject dimensions in points Aspose.Cells example
// Tags: Aspose.Cells chartobject width height | set chart dimensions points C# | resize column chart Aspose.Cells | chart layout consistency Excel .NET | programmatic chart sizing Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartResizeDemo
{
    // Creates a workbook, adds sample data, inserts a column chart, resizes it to 400 × 300 points, and saves the file as ResizedChart.xlsx.
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
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
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
