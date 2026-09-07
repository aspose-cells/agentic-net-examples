// Title: Programmatically toggle the legend visibility of a column chart in Aspose.Cells using C#
// AI Prompts: Generate a C# method that creates a workbook, adds a column chart, and sets its ShowLegend property based on a boolean parameter. | Write code to conditionally display or hide the legend of an Aspose.Cells chart according to a user‑supplied flag, then save the workbook as an .xlsx file. | Adapt an existing Aspose.Cells chart example to accept a bool argument that controls chart.ShowLegend before exporting.
// Common Searches: Aspose.Cells C# how to hide chart legend based on a variable | set ShowLegend property of a chart dynamically in Aspose.Cells | toggle column chart legend visibility with a boolean in C# Aspose.Cells | programmatically control chart legend display in Aspose.Cells workbook | C# Aspose.Cells example for showing or hiding chart legend
// Tags: Aspose.Cells chart ShowLegend | C# toggle chart legend | column chart legend visibility Aspose.Cells | programmatic chart legend control Aspose.Cells | save workbook with chart legend Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a new workbook, fills it with sample data, adds a column chart, and sets the chart's ShowLegend property according to a supplied boolean flag before saving the file as ToggleLegendVisibility.xlsx.
    public class ToggleLegendVisibility
    {
        /// <param name="showLegend">If true the legend will be displayed; otherwise it will be hidden.</param>
        public static void Run(bool showLegend)
        {
            try
            {
                // Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Item");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["A3"].PutValue("Orange");
                sheet.Cells["A4"].PutValue("Banana");
                sheet.Cells["B1"].PutValue("Quantity");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Toggle legend visibility based on the flag
                chart.ShowLegend = showLegend;
                Console.WriteLine($"Legend visibility set to: {chart.ShowLegend}");

                // Save the workbook (lifecycle: save)
                string outputPath = "ToggleLegendVisibility.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            // Example usage: show the legend
            ToggleLegendVisibility.Run(true);
        }
    }
}
