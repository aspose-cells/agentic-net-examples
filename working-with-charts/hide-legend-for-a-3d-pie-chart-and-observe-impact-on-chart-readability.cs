// Title: Hide Legend in a 3‑D Pie Chart using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add fruit data, generate a 3‑D pie chart, disable its legend with the ShowLegend property, set a title, and save the file as XLSX. The example also discusses how removing the legend affects chart readability.
// Keywords: Aspose.Cells hide legend | C# 3D pie chart legend | Chart.ShowLegend false | remove legend Aspose.Cells | Excel chart legend visibility .NET | Aspose.Cells Pie3D example | disable chart legend programmatically
// Common Searches: Aspose.Cells hide legend C# | turn off legend in 3D pie chart .NET | Chart.ShowLegend property example | remove legend from Excel chart using Aspose | how to hide chart legend in C#
// Developer Intent: Programmatically hide the legend of a 3‑D pie chart in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Produce a compact sales distribution chart where space is limited and the legend is unnecessary. | Create a presentation slide that relies on slice colors and a descriptive title, eliminating visual clutter. | Generate automated reports that embed multiple charts and need a cleaner layout without redundant legends.
// AI Prompts: Show C# code that creates a 3‑D pie chart with Aspose.Cells and hides its legend while keeping the title. | Explain the effect of disabling the legend on a 3‑D pie chart’s readability and suggest alternative labeling techniques. | Provide a step‑by‑step guide to use Chart.ShowLegend = false in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add fruit data, generate a 3‑D pie chart, disable its legend with the ShowLegend property, set a title, and save the file as XLSX. The example also discusses how removing the legend affects chart readability.
    public class HideLegend3DPieChart
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pie chart
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

                // Hide the legend
                chart.ShowLegend = false;

                // Optional: give the chart a title for context
                chart.Title.Text = "Fruit Distribution (Legend Hidden)";

                // Save the workbook to an XLSX file
                string outputPath = "HideLegend3DPieChart.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            HideLegend3DPieChart.Run();
        }
    }
}
