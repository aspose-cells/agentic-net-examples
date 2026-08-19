// Title: C# – Create Column Chart with Bottom Legend Shifted Up Using Aspose.Cells
// Description: Demonstrates how to generate a workbook, add sample data, insert a column chart, place the legend at the bottom, raise it by a 5 % vertical offset with the YRatioToChart property, recalculate the layout, and save the file.
// Keywords: Aspose.Cells chart legend bottom | C# shift legend upward | YRatioToChart offset | column chart legend position | Aspose.Cells legend customization
// Common Searches: Aspose.Cells move legend up after setting bottom | C# chart legend offset percentage Aspose.Cells | How to adjust legend YRatioToChart in .NET | Shift chart legend upward in Aspose.Cells
// Developer Intent: Create a column chart, set its legend to the bottom, then nudge the legend upward by a fixed percentage of the chart height using Aspose.Cells for .NET.
// Use Cases: Financial reports where the legend must stay near the bottom but not overlap the chart title. | Automated dashboards that require precise vertical spacing of legends for visual balance. | Presentation‑ready Excel files with legends anchored at the bottom and slightly raised for layout consistency.
// AI Prompts: Show C# code to place a chart legend at the bottom and move it up by a specific percentage with Aspose.Cells. | Provide an example of using the YRatioToChart property to offset a legend after positioning it at the bottom in Aspose.Cells for .NET. | Explain how to calculate a dynamic vertical offset for a chart legend based on chart height in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;   // Required for Chart, ChartType, LegendPositionType

namespace AsposeCellsExample
{
    // Demonstrates how to generate a workbook, add sample data, insert a column chart, place the legend at the bottom, raise it by a 5 % vertical offset with the YRatioToChart property, recalculate the layout, and save the file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add sample data for the chart
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

                // Set the data range for the chart (including headers)
                chart.SetChartDataRange("A1:B4", true);

                // Position the legend at the bottom of the chart
                chart.Legend.Position = LegendPositionType.Bottom;

                // Shift the legend upward by a fixed offset (5% of the chart height)
                double offset = 0.05; // 5% upward
                chart.Legend.YRatioToChart = Math.Max(0, chart.Legend.YRatioToChart - offset);

                // Recalculate the chart layout to apply changes
                chart.Calculate();

                // Save the workbook
                string outputPath = "ChartWithShiftedLegend.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the chart:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
