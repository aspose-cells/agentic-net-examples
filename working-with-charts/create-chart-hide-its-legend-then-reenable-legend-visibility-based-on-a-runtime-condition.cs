// Title: Hide an Excel chart legend with Aspose.Cells in C# and show it conditionally at runtime
// AI Prompts: Write C# code that uses Aspose.Cells to create a column chart, set ShowLegend = false, then set ShowLegend = true only when a custom boolean method returns true, and save the workbook. | Generate a method that populates worksheet data, adds a chart, and toggles the chart's legend visibility based on a runtime condition using Aspose.Cells APIs. | Provide a complete Aspose.Cells example that demonstrates initially hiding the chart legend, evaluating a runtime condition, and re‑enabling the legend before exporting to an .xlsx file.
// Common Searches: aspocells c# hide chart legend then show based on condition | how to programmatically toggle Excel chart legend visibility with Aspose.Cells | C# Aspose.Cells example for conditional chart legend display | runtime check to enable chart legend in Aspose.Cells workbook | Aspose.Cells ShowLegend property usage in C#
// Tags: Aspose.Cells chart legend visibility | C# hide Excel chart legend programmatically | conditional chart legend toggle Aspose.Cells | column chart creation Aspose.Cells C# | export workbook with dynamic legend Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartLegendToggle
{
    // The example creates a workbook, adds sample data, inserts a column chart, hides its legend, evaluates a runtime condition, re‑enables the legend if the condition is true, and saves the file as ChartLegendToggle.xlsx.
    class Program
    {
        static void Main(string[] args)
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

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Hide the legend initially
            chart.ShowLegend = false;
            Console.WriteLine("Legend hidden: " + chart.ShowLegend);

            // Runtime condition that determines whether the legend should be shown
            bool shouldShowLegend = GetRuntimeCondition();

            // Re‑enable legend visibility based on the condition
            if (shouldShowLegend)
            {
                chart.ShowLegend = true;
                Console.WriteLine("Condition met – legend re‑enabled: " + chart.ShowLegend);
            }
            else
            {
                Console.WriteLine("Condition not met – legend remains hidden.");
            }

            // Save the workbook to a file
            workbook.Save("ChartLegendToggle.xlsx");
        }

        // Example method that simulates a runtime condition (replace with real logic)
        static bool GetRuntimeCondition()
        {
            // For demonstration, toggle based on current second being even
            return DateTime.Now.Second % 2 == 0;
        }
    }
}
