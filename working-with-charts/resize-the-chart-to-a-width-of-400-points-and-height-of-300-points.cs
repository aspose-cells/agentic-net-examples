// Title: Resize an Aspose.Cells chart in C# – set width to 400 pt and height to 300 pt
// Description: This example creates a workbook, adds sample data, inserts a column chart, and then programmatically sets the chart size by assigning 400 points to ChartObject.Width and 300 points to ChartObject.Height before saving as ResizedChart.xlsx.
// Keywords: Aspose.Cells | C# | .NET | chart resize | ChartObject.Width | ChartObject.Height | set chart dimensions | Excel chart size | column chart | programmatic chart sizing | Excel automation | point units
// Common Searches: Aspose.Cells set chart width and height C# | Resize Excel chart to 400 points by 300 points using Aspose | How to change chart size programmatically in .NET | ChartObject.Width and Height examples Aspose.Cells | Adjust Excel chart dimensions with C# code
// Developer Intent: Set the chart's width to 400 pt and height to 300 pt using Aspose.Cells in a .NET application.
// Use Cases: Create a column chart with exact dimensions for a standardized report layout. | Generate multiple charts with consistent sizing across automated Excel files. | Fit a chart into a predefined page or dashboard area when building Excel dashboards programmatically.
// AI Prompts: Show C# code to resize an Aspose.Cells chart to 400 pt × 300 pt. | Explain how ChartObject.Width and ChartObject.Height affect the rendered size of an Excel chart. | Provide a step‑by‑step guide for setting chart dimensions in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This example creates a workbook, adds sample data, inserts a column chart, and then programmatically sets the chart size by assigning 400 points to ChartObject.Width and 300 points to ChartObject.Height before saving as ResizedChart.xlsx.
class ResizeChartExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data for the chart
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
