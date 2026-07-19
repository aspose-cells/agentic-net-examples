// Title: Hide Chart Legends in Multiple Column Charts Using Aspose.Cells for .NET
// Description: Shows how to build a workbook, populate sample data, create three column charts, disable each chart's legend with the ShowLegend property, assign titles, and save the file as ChartsWithoutLegend.xlsx.
// Keywords: Aspose.Cells C# hide chart legend | ShowLegend false Aspose.Cells | remove legend from Excel chart programmatically | multiple charts without legend .NET | column chart legend suppression | Aspose.Cells chart formatting | Excel chart legend visibility | global
// Common Searches: Aspose.Cells hide legend C# | turn off chart legend in Aspose.Cells | create charts without legends Aspose.Cells .NET | batch hide legends in Excel charts programmatically | Aspose.Cells chart legend visibility
// Developer Intent: The developer needs to generate several charts and suppress their legends to keep the visual layout clean for presentations or reports.
// Use Cases: Build a slide‑deck workbook with three column charts that have no legends, reducing visual clutter. | Produce a compact PDF or printed report where chart legends are unnecessary and waste space. | Export charts to PowerPoint where overlapping legends would interfere with slide design.
// AI Prompts: Provide C# code that disables the legend for every chart in an existing Aspose.Cells workbook. | Show how to set the ShowLegend property conditionally based on runtime logic in Aspose.Cells. | Explain how to programmatically remove legends from a batch of charts while keeping other formatting intact.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartLegendHide
{
    // Shows how to build a workbook, populate sample data, create three column charts, disable each chart's legend with the ShowLegend property, assign titles, and save the file as ChartsWithoutLegend.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data for the charts
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Generate a series of charts (e.g., three column charts) positioned differently
            for (int i = 0; i < 3; i++)
            {
                // Add a chart of type Column
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5 + i * 15, 0, 15 + i * 15, 5);
                Chart chart = sheet.Charts[chartIndex];

                // Set data range for the chart
                chart.NSeries.Add("B2:C4", true);          // Values from both series
                chart.NSeries.CategoryData = "A2:A4";      // Categories

                // Hide the legend to reduce visual noise
                chart.ShowLegend = false;

                // Optional: give each chart a title to differentiate
                chart.Title.Text = $"Chart {i + 1}";
            }

            // Save the workbook containing the charts
            workbook.Save("ChartsWithoutLegend.xlsx");
        }
    }
}
