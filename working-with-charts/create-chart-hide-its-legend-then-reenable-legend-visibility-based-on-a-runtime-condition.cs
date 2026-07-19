// Title: Aspose.Cells for .NET – Toggle Chart Legend Visibility at Runtime (C#)
// Description: This example demonstrates how to create a column chart with Aspose.Cells, hide its legend using the ShowLegend property, and then re‑enable the legend when a runtime boolean condition is met, before saving the workbook as an XLSX file.
// Keywords: Aspose.Cells | C# chart legend | ShowLegend property | hide chart legend | runtime toggle legend | conditional chart legend | Excel chart Aspose.Cells | dynamic legend visibility | column chart example | programmatic chart styling
// Common Searches: Aspose.Cells hide legend then show based on condition | C# toggle chart legend at runtime | How to use ShowLegend in Aspose.Cells | Conditional chart legend visibility .NET | Programmatically control Excel chart legend with Aspose
// Developer Intent: The developer needs to create a chart, suppress its legend initially, and display the legend only when a specific runtime condition evaluates to true.
// Use Cases: Generate financial reports where legends appear only for charts that exceed a threshold value. | Build a compact Excel dashboard that reveals legends when a user expands a chart section. | Export data to Excel with a configuration flag that determines whether chart legends are shown.
// AI Prompts: Write C# code with Aspose.Cells to add a column chart, hide its legend, and show the legend when a boolean variable is true. | Provide an Aspose.Cells example that reads a setting from appsettings.json and toggles the ShowLegend property accordingly. | Explain step‑by‑step how to control chart legend visibility in Aspose.Cells for .NET based on a runtime condition.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartLegendToggle
{
    // This example demonstrates how to create a column chart with Aspose.Cells, hide its legend using the ShowLegend property, and then re‑enable the legend when a runtime boolean condition is met, before saving the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Runtime condition – for demo purposes we use a simple boolean.
            // In real scenarios this could be based on user input, configuration, etc.
            bool shouldShowLegend = GetRuntimeCondition();

            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart.
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart.
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart.
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Hide the legend initially.
            chart.ShowLegend = false;
            Console.WriteLine("Legend hidden: " + chart.ShowLegend);

            // Re‑enable legend visibility based on the runtime condition.
            if (shouldShowLegend)
            {
                chart.ShowLegend = true;
                Console.WriteLine("Runtime condition met – legend shown: " + chart.ShowLegend);
            }
            else
            {
                Console.WriteLine("Runtime condition not met – legend remains hidden.");
            }

            // Save the workbook.
            workbook.Save("ChartLegendToggle.xlsx");
        }

        // Example method to determine the runtime condition.
        // Replace with actual logic as needed.
        static bool GetRuntimeCondition()
        {
            // For illustration, toggle based on the current second being even.
            return DateTime.Now.Second % 2 == 0;
        }
    }
}
