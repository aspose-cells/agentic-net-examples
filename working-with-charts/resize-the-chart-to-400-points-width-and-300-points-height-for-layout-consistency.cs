// Title: C# – Resize an Aspose.Cells chart to 400 pt × 300 pt for consistent layout
// Description: Creates a workbook, adds a column chart with sample data, then sets ChartObject.Width = 400 and ChartObject.Height = 300 points before saving the file as ResizedChart.xlsx.
// Keywords: Aspose.Cells | C# | chart resize | set chart width | set chart height | ChartObject dimensions | Excel chart size points | layout consistency | programmatic chart sizing
// Common Searches: Aspose.Cells set chart size points | Resize chart to 400x300 Aspose.Cells C# | Change Excel chart dimensions programmatically | How to adjust ChartObject width and height in Aspose.Cells | Set chart object size in points using Aspose.Cells .NET
// Developer Intent: Set the chart’s width to 400 points and height to 300 points to achieve a uniform appearance in the generated workbook.
// Use Cases: Standardize chart dimensions across multiple worksheets in automated reports. | Fit charts into predefined placeholders when building dashboard workbooks. | Ensure exported charts match exact size requirements for branding in PDFs or images.
// AI Prompts: Generate C# code that resizes an Aspose.Cells chart to 500 pt × 400 pt and centers it on the worksheet. | Show how to adjust chart dimensions dynamically based on page size using Aspose.Cells. | Explain how to retrieve and modify ChartObject Width and Height properties in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartResize
{
    // Creates a workbook, adds a column chart with sample data, then sets ChartObject.Width = 400 and ChartObject.Height = 300 points before saving the file as ResizedChart.xlsx.
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
