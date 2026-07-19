// Title: Create a transparent background series for a stacked column progress bar with Aspose.Cells for .NET
// Description: Demonstrates how to add a stacked column chart in an Excel workbook, insert a visible progress series and an invisible background series, set the background series area transparency to 100 %, hide it from the legend, and save the file as an .xlsx document using Aspose.Cells for .NET.
// Keywords: Aspose.Cells transparent series | stacked column progress bar .NET | hide chart series legend Aspose | Excel chart area transparency | simulate progress bar Aspose.Cells
// Common Searches: Aspose.Cells make chart series transparent | progress bar chart using stacked column in C# | remove series from legend Aspose.Cells | set series area transparency Excel chart | create progress bar effect with Aspose.Cells
// Developer Intent: Apply full transparency to the second series of a stacked column chart so it acts as an invisible background for a progress‑bar visual.
// Use Cases: Dashboard that shows task completion as compact progress bars without extra legend entries. | Printable project status reports where each row contains a clean, self‑contained progress indicator. | Automated Excel generation for milestone tracking that uses a hidden series to define bar length while keeping the chart uncluttered.
// AI Prompts: Generate C# code with Aspose.Cells that builds a stacked column chart, makes the background series fully transparent, and excludes it from the legend. | Explain how setting Series.Area.Transparency to 1.0 hides a chart series and why this technique is useful for progress‑bar visualizations. | Modify the example to use a separate data range for the invisible series representing total capacity instead of duplicating the progress values.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace ProgressBarChartDemo
{
    // Demonstrates how to add a stacked column chart in an Excel workbook, insert a visible progress series and an invisible background series, set the background series area transparency to 100 %, hide it from the legend, and save the file as an .xlsx document using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data for the progress bar
            // Column A – categories, Column B – actual progress values
            sheet.Cells["A1"].PutValue("Task");
            sheet.Cells["A2"].PutValue("Task 1");
            sheet.Cells["A3"].PutValue("Task 2");
            sheet.Cells["A4"].PutValue("Task 3");

            sheet.Cells["B1"].PutValue("Progress");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(60);
            sheet.Cells["B4"].PutValue(90);

            // Add a stacked column chart (used to simulate a progress bar)
            int chartIndex = sheet.Charts.Add(ChartType.ColumnStacked, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // First series – actual progress (visible)
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Second series – invisible background (used to create the bar background)
            // Use the same range for simplicity; it will be made fully transparent
            chart.NSeries.Add("B2:B4", true);
            // Access the second series (index 1)
            Series invisibleSeries = chart.NSeries[1];

            // Make the invisible series fully transparent
            // Transparency value 1.0 means completely clear
            invisibleSeries.Area.Transparency = 1.0;

            // Optionally hide the series from the legend
            invisibleSeries.IsFiltered = true;

            // Save the workbook
            workbook.Save("ProgressBarChart.xlsx", SaveFormat.Xlsx);
        }
    }
}
