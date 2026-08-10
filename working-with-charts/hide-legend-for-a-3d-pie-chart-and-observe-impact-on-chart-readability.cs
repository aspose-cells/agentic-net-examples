// Title: Hide Legend in a 3‑D Pie Chart Using Aspose.Cells for .NET (C#)
// Description: Creates a workbook, fills fruit data, adds a 3‑D pie chart, disables its legend with chart.ShowLegend = false, sets a title, and saves the file. Demonstrates how removing the legend affects chart readability.
// Keywords: Aspose.Cells hide legend | 3D pie chart C# | chart.ShowLegend property | remove legend Excel chart Aspose | chart readability Aspose.Cells | .NET chart customization
// Common Searches: how to hide legend in Aspose.Cells chart | disable legend 3D pie chart C# | Aspose.Cells ShowLegend false example | impact of legend removal on chart readability | C# code to hide Excel chart legend
// Developer Intent: Programmatically hide a chart legend to evaluate visual clarity.
// Use Cases: Produce a dashboard where the legend duplicates data labels, reducing visual clutter. | Generate multiple 3‑D pie charts in a single report and hide legends for a compact layout. | Create a toggle feature that switches legend visibility on‑demand to compare readability.
// AI Prompts: Write C# code with Aspose.Cells that builds a 3‑D pie chart, hides its legend, and saves the workbook. | Explain the effect of the ShowLegend property on chart rendering and when it is advisable to hide the legend. | Suggest alternative approaches for controlling legend visibility across different chart types in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendDemo
{
    // Creates a workbook, fills fruit data, adds a 3‑D pie chart, disables its legend with chart.ShowLegend = false, sets a title, and saves the file. Demonstrates how removing the legend affects chart readability.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the 3‑D pie chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["A4"].PutValue("Banana");
            sheet.Cells["A5"].PutValue("Grape");

            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(25);
            sheet.Cells["B5"].PutValue(15);

            // Add a 3‑D pie chart (ChartType.Pie3D)
            int chartIndex = sheet.Charts.Add(ChartType.Pie3D, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // By default the legend is visible; hide it to see the effect on readability
            chart.ShowLegend = false;

            // Optional: give the chart a title for context
            chart.Title.Text = "Fruit Distribution (3‑D Pie)";

            // Save the workbook
            workbook.Save("3DPieChart_HideLegend.xlsx");
        }
    }
}
