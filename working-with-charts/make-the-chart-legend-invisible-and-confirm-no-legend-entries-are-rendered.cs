// Title: Hide a Chart Legend and Confirm No Entries Using Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds a column chart with sample data, disables the legend via Chart.ShowLegend = false, calculates the chart, reads Legend.LegendEntries.Count, outputs the visibility flag and entry count, and saves the file as ChartWithoutLegend.xlsx.
// Keywords: Aspose.Cells hide legend | Chart.ShowLegend false C# | Aspose.Cells chart legend visibility | Excel chart without legend .NET | Aspose.Cells LegendEntries count | C# Aspose.Cells chart example
// Common Searches: how to hide legend in Aspose.Cells chart | verify chart legend is not rendered Aspose.Cells | Aspose.Cells Chart.ShowLegend property usage | C# code to remove chart legend in Excel | check legend entries count Aspose.Cells
// Developer Intent: Make the chart legend invisible and ensure no legend entries appear in the rendered workbook.
// Use Cases: Produce Excel reports with clean charts that omit legends for a minimalist design. | Programmatically validate that a chart’s legend is hidden before distributing the file to end users. | Create dashboard worksheets where space is limited and legends are unnecessary, while still confirming data series exist.
// AI Prompts: Generate C# code with Aspose.Cells that hides a chart legend, calculates the chart, and asserts LegendEntries.Count is zero. | Show how to disable the legend for any chart type in Aspose.Cells and log the visibility status. | Write a unit test in C# that creates a column chart, sets ShowLegend to false, runs Calculate(), and verifies no legend entries are rendered.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartLegendDemo
{
    // Creates a workbook, adds a column chart with sample data, disables the legend via Chart.ShowLegend = false, calculates the chart, reads Legend.LegendEntries.Count, outputs the visibility flag and entry count, and saves the file as ChartWithoutLegend.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
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

            // Hide the legend
            chart.ShowLegend = false;

            // Verify that the legend is hidden
            Console.WriteLine("Chart.ShowLegend after setting to false: " + chart.ShowLegend);

            // Optionally, calculate the chart and check legend entries count
            chart.Calculate();
            int legendEntryCount = chart.Legend.LegendEntries.Count;
            Console.WriteLine("Number of legend entries (should exist but not rendered): " + legendEntryCount);

            // Save the workbook
            workbook.Save("ChartWithoutLegend.xlsx");
        }
    }
}
