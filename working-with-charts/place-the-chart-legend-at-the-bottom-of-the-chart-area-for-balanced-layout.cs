// Title: How to place a chart legend at the bottom of a column chart with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells to create a column chart and set its legend position to the bottom of the chart area. | Show how to disable automatic legend sizing and assign a specific width and height to the legend in an Aspose.Cells chart. | Demonstrate saving the workbook after configuring legend placement and custom dimensions with Aspose.Cells.
// Common Searches: aspnet aspocells set chart legend bottom position c# example | c# Aspose.Cells column chart legend placement at bottom of chart area | disable automatic legend size and set custom dimensions in Aspose.Cells chart | how to move Excel chart legend to bottom using Aspose.Cells .NET | Aspose.Cells legend position bottom column chart code sample
// Tags: Aspose.Cells chart legend positioning | C# column chart legend bottom | Aspose.Cells custom legend size | Excel chart legend placement Aspose.Cells | Aspose.Cells set legend dimensions

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendBottomExample
{
    // Creates a new workbook, adds sample data, generates a column chart, moves the legend to the bottom, disables automatic sizing, sets a custom width and height for the legend, and saves the file as ChartWithBottomLegend.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
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

            // Position the legend at the bottom of the chart area
            chart.Legend.Position = LegendPositionType.Bottom;

            // Optional: adjust legend size for better appearance
            chart.Legend.IsAutomaticSize = false;
            chart.Legend.Width = 400;
            chart.Legend.Height = 50;

            // Save the workbook
            workbook.Save("ChartWithBottomLegend.xlsx");
        }
    }
}
