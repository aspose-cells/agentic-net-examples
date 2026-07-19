// Title: Toggle Chart Legend Visibility in Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds sample data and a column chart, then sets the chart's ShowLegend property based on a Boolean flag before saving as ChartLegendToggled.xlsx.
// Keywords: Aspose.Cells | C# chart legend | ShowLegend property | toggle legend visibility | hide chart legend programmatically | column chart Aspose.Cells | .NET workbook automation | conditional chart formatting
// Common Searches: Aspose.Cells hide legend C# | ShowLegend property example .NET | toggle chart legend with boolean | programmatically control chart legend Aspose | C# set chart legend visibility
// Developer Intent: Control whether a chart legend is displayed by setting Chart.ShowLegend based on a Boolean parameter.
// Use Cases: Hide legends for single‑series charts while showing them for multi‑series reports. | Expose a configuration option that lets end‑users decide if a legend should appear. | Integrate external settings (e.g., JSON, database flag) to conditionally show or hide legends during automated workbook generation.
// AI Prompts: Generate a C# method that accepts a bool and applies it to chart.ShowLegend in Aspose.Cells, then saves the file. | Show how to read a configuration value and use it to set the legend visibility when building a chart with Aspose.Cells. | Explain the steps to conditionally hide a chart legend for single‑series data using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendToggle
{
    // Creates a workbook, adds sample data and a column chart, then sets the chart's ShowLegend property based on a Boolean flag before saving as ChartLegendToggled.xlsx.
    public class LegendToggleDemo
    {
        // Toggles the chart legend visibility based on the supplied flag.
        public static void Run(bool showLegend)
        {
            // Create a new workbook (lifecycle rule: create)
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
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Apply the legend visibility flag (rule: Chart.ShowLegend)
            chart.ShowLegend = showLegend;
            Console.WriteLine($"Legend visibility set to: {chart.ShowLegend}");

            // Save the workbook (lifecycle rule: save)
            workbook.Save("ChartLegendToggled.xlsx");
        }

        // Example entry point
        public static void Main()
        {
            // Example: hide the legend
            Run(false);

            // Example: show the legend
            Run(true);
        }
    }
}
