// Title: Aspose.Cells C# – Hide Chart Legend and Reveal It Conditionally
// Description: The sample builds a workbook, inserts a column chart with sample data, disables the legend, evaluates a runtime boolean, and re‑enables the legend when the condition is met before saving as ChartLegendConditional.xlsx.
// Keywords: Aspose.Cells chart legend hide | C# conditional legend display | Aspose.Cells runtime chart settings | toggle legend Aspose.Cells | Excel chart legend C# example | GitHub Aspose.Cells chart demo | global Aspose.Cells chart manipulation
// Common Searches: how to hide a chart legend using Aspose.Cells .NET | show chart legend only when a condition is true C# | Aspose.Cells hide and show legend at runtime | conditional chart legend visibility Aspose.Cells | example of chart legend toggle in C#
// Developer Intent: Suppress a chart legend initially and display it later only if a runtime condition evaluates to true.
// Use Cases: Create a clean‑look report where the legend appears only for large data sets. | Build an interactive dashboard that reveals the legend based on user selections. | Generate Excel exports that toggle legend visibility according to a configuration flag.
// AI Prompts: Generate C# code with Aspose.Cells that adds a line chart, hides its legend, and shows it when a boolean variable is true. | Provide an Aspose.Cells snippet that changes chart legend visibility based on a worksheet cell value. | Write a method that accepts a Chart object and a predicate, then sets ShowLegend accordingly.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartLegendDemo
{
    // The sample builds a workbook, inserts a column chart with sample data, disables the legend, evaluates a runtime boolean, and re‑enables the legend when the condition is met before saving as ChartLegendConditional.xlsx.
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

            // Initially hide the legend
            chart.ShowLegend = false;
            Console.WriteLine("Legend hidden: " + chart.ShowLegend);

            // Runtime condition to decide whether to show the legend again
            bool shouldShowLegend = GetRuntimeCondition();

            // Re‑enable legend visibility based on the condition
            if (shouldShowLegend)
            {
                chart.ShowLegend = true;
                Console.WriteLine("Legend re‑enabled: " + chart.ShowLegend);
            }

            // Save the workbook
            workbook.Save("ChartLegendConditional.xlsx");
        }

        // Example method that determines the runtime condition.
        // Replace with actual logic as needed.
        static bool GetRuntimeCondition()
        {
            // For demonstration, toggle based on current second (even = true, odd = false)
            return DateTime.Now.Second % 2 == 0;
        }
    }
}
