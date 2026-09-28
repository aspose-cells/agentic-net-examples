// Title: Move chart legend to the bottom and hide its border using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that positions a chart legend at the bottom and turns off its border line. | Show how to adjust a column chart's legend placement and hide the legend outline in Aspose.Cells for .NET. | Create a workbook, add data, generate a column chart, set the legend to the bottom, and disable the legend border using Aspose.Cells.
// Common Searches: Aspose.Cells C# place chart legend at bottom | how to hide legend border Aspose.Cells chart | set legend position bottom Aspose.Cells .NET example | remove legend line from chart using Aspose.Cells C#
// Tags: Aspose.Cells set legend position bottom | Aspose.Cells hide legend border | C# column chart legend formatting Aspose.Cells | Aspose.Cells legend line visibility control | Aspose.Cells workbook save chart layout

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendExample
{
    // // Demonstrates creating a workbook, adding sample data, inserting a column chart, moving its legend to the bottom, and hiding the legend border using Aspose.Cells for .NET.
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

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Move the legend to the bottom of the chart
            chart.Legend.Position = LegendPositionType.Bottom;

            // Hide the legend border for a cleaner layout
            // The Border property returns a Line object; setting IsVisible to false hides it
            chart.Legend.Border.IsVisible = false;

            // Save the workbook
            workbook.Save("ChartWithBottomLegend.xlsx");
        }
    }
}
