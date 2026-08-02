// Title: Hide Chart Legend in Aspose.Cells for .NET – Minimalist Column Chart Example
// Description: Shows how to create a workbook, add a column chart, and hide its legend with Aspose.Cells for .NET (C#) to produce a clean, minimalist chart suitable for presentations.
// Keywords: Aspose.Cells | C# | hide chart legend | ShowLegend false | minimalist chart | column chart | Excel chart formatting | Aspose.Cells chart API | remove legend | presentation slide chart
// Common Searches: Aspose.Cells hide legend C# | remove chart legend Aspose.Cells .NET | minimalist chart Aspose.Cells example | disable legend in Aspose.Cells chart | Aspose.Cells ShowLegend property usage
// Developer Intent: Disable the chart legend to create a minimalist chart layout in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Design clean charts for PowerPoint slides by omitting legends. | Generate reports where legends add unnecessary visual noise. | Build dashboards with a single shared legend, hiding individual ones. | Create printable Excel charts with reduced clutter.
// AI Prompts: Generate C# code with Aspose.Cells to create a bar chart and hide its legend. | Show how to toggle the ShowLegend property based on a boolean parameter in Aspose.Cells. | Explain how to hide legends for all charts in a workbook using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace HideChartLegendDemo
{
    // Shows how to create a workbook, add a column chart, and hide its legend with Aspose.Cells for .NET (C#) to produce a clean, minimalist chart suitable for presentations.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
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
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Hide the legend to achieve a minimalist layout
            chart.ShowLegend = false;

            // Save the workbook to a file
            workbook.Save("MinimalistChart.xlsx");
        }
    }
}
