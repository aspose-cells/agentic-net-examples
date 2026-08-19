// Title: Toggle Excel Chart Legend Visibility with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add a column chart, and control the chart's legend using the ShowLegend property driven by a Boolean flag, then save the file as an .xlsx document.
// Keywords: Aspose.Cells chart legend toggle | C# ShowLegend property | programmatically hide Excel chart legend | Aspose.Cells .NET chart customization | dynamic legend visibility Excel
// Common Searches: Aspose.Cells hide chart legend C# | set ShowLegend flag Aspose.Cells | toggle Excel chart legend programmatically | how to control chart legend visibility with Aspose.Cells | C# example for chart legend visibility
// Developer Intent: Enable developers to turn a chart legend on or off in an Excel file based on a runtime Boolean value.
// Use Cases: Create financial dashboards where the legend appears only for multi‑series charts. | Provide end‑users a checkbox to show or hide legends in generated reports. | Automate report generation that suppresses legends for single‑category charts to conserve space.
// AI Prompts: Generate C# code using Aspose.Cells that adds a line chart and hides its legend when a variable hideLegend is true. | Show how to read a configuration setting and apply chart.ShowLegend for an existing workbook in Aspose.Cells. | Explain a method to iterate through all charts in a workbook and toggle their legends without recreating the charts.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add a column chart, and control the chart's legend using the ShowLegend property driven by a Boolean flag, then save the file as an .xlsx document.
    public class ToggleLegendVisibilityDemo
    {
        // Toggles the chart legend visibility based on the supplied flag.
        public static void Run(bool showLegend)
        {
            try
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
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Apply the user‑defined legend visibility flag
                chart.ShowLegend = showLegend;
                Console.WriteLine($"Legend visibility set to: {chart.ShowLegend}");

                // Save the workbook
                workbook.Save("ToggleLegendVisibility.xlsx");
                Console.WriteLine("Workbook saved as ToggleLegendVisibility.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            // Example usage: show the legend
            ToggleLegendVisibilityDemo.Run(true);
        }
    }
}
