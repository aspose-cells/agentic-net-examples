// Title: Aspose.Cells C# – Create a Column Chart, Hide Its Legend, and Show It Conditionally at Runtime
// Description: Demonstrates how to build a workbook with a column chart using Aspose.Cells, hide the chart legend via the ShowLegend property, and re‑enable the legend when a runtime condition (e.g., a command‑line argument) is true, then save the file.
// Keywords: Aspose.Cells chart legend hide | Aspose.Cells ShowLegend property | C# toggle chart legend | runtime chart legend Aspose.Cells | conditional legend visibility | create column chart Aspose.Cells | programmatic chart legend control
// Common Searches: how to hide chart legend Aspose.Cells C# | show chart legend based on condition Aspose.Cells | Aspose.Cells toggle legend visibility at runtime | C# example for conditional chart legend | Aspose.Cells command line argument legend
// Developer Intent: Programmatically control a chart's legend visibility in an Aspose.Cells workbook, hiding it by default and displaying it only when a specified runtime condition is met.
// Use Cases: Generate compact Excel reports where legends are omitted for space, but include them when a user requests a detailed view. | Create automated dashboards that suppress legends on mobile exports and reveal them on desktop exports via a flag. | Produce multiple workbooks where legend visibility is driven by configuration settings such as command‑line arguments or JSON files.
// AI Prompts: Write C# code with Aspose.Cells to add a line chart, hide its legend, and enable the legend only when a boolean variable `showLegend` is true. | Provide an Aspose.Cells example that reads a setting from appsettings.json and toggles the chart legend accordingly. | Explain how to use the ShowLegend property to hide and later show a chart legend after adding series in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartLegendToggle
{
    // Demonstrates how to build a workbook with a column chart using Aspose.Cells, hide the chart legend via the ShowLegend property, and re‑enable the legend when a runtime condition (e.g., a command‑line argument) is true, then save the file.
    class Program
    {
        static void Main(string[] args)
        {
            // Runtime condition: show legend if first argument is "true"
            bool enableLegend = args.Length > 0 && args[0].Equals("true", StringComparison.OrdinalIgnoreCase);

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

            // Initially hide the legend
            chart.ShowLegend = false;

            // Re‑enable legend visibility based on the runtime condition
            if (enableLegend)
            {
                chart.ShowLegend = true;
            }

            // Save the workbook
            workbook.Save("ChartLegendToggle.xlsx");
        }
    }
}
