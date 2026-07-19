// Title: Hide Chart Legend in Aspose.Cells for .NET – Minimalist Column Chart Example
// Description: C# code that creates a workbook, adds sample data, inserts a column chart, disables the legend with chart.ShowLegend = false, and saves the file as ChartWithoutLegend.xlsx, delivering a clean chart ready for presentation slides.
// Keywords: Aspose.Cells hide legend | remove chart legend C# | ShowLegend false | minimalist chart Aspose.Cells | column chart without legend | Excel export C# | presentation slide chart | Aspose.Cells chart example
// Common Searches: how to hide legend in Aspose.Cells chart C# | Aspose.Cells remove chart legend .NET | minimalist chart example Aspose.Cells | ShowLegend property false Aspose.Cells | create column chart without legend using Aspose
// Developer Intent: Remove the legend from an Excel chart via Aspose.Cells for .NET to achieve a simplified visual layout.
// Use Cases: Generating slide‑ready charts that do not require a legend. | Building dashboards where series names are displayed directly on the chart. | Exporting Excel reports with streamlined graphics for concise communication.
// AI Prompts: Write C# code with Aspose.Cells to hide the legend of a pie chart and resize the chart. | Explain how to toggle chart legend visibility at runtime in a .NET application using Aspose.Cells. | Provide a loop that hides legends for all charts in a workbook with Aspose.Cells. | Show how to move a chart legend to a custom position or hide it based on a condition.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendHideDemo
{
    // C# code that creates a workbook, adds sample data, inserts a column chart, disables the legend with chart.ShowLegend = false, and saves the file as ChartWithoutLegend.xlsx, delivering a clean chart ready for presentation slides.
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

            // Hide the chart legend for a minimalist layout
            chart.ShowLegend = false;

            // Save the workbook to a file
            workbook.Save("ChartWithoutLegend.xlsx");
        }
    }
}
