// Title: Bind Worksheet Column Values to a Stacked Bar Progress Bar Chart with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, fills column A with task names and column B with numeric progress percentages, adds a stacked bar chart, binds the B2:B10 range to the first series, assigns A2:A10 as category labels, disables series filtering, and sets Overlap = -100 and GapWidth = 0 to render a classic progress‑bar look before saving as an XLSX file.
// Keywords: Aspose.Cells bind column to chart series | C# progress bar chart Aspose.Cells | stacked bar chart progress percentages .NET | chart series visibility Aspose.Cells | Overlap GapWidth progress bar styling
// Common Searches: Aspose.Cells bind column B to stacked bar series | Create progress bar chart from worksheet data in C# | Set IsFiltered false for chart series Aspose.Cells | Adjust Overlap and GapWidth for progress bar appearance | Generate Excel progress bar chart programmatically
// Developer Intent: Generate a stacked bar chart that visualizes worksheet column values as visible progress bars using Aspose.Cells for .NET.
// Use Cases: Project status reports showing each task’s completion as a compact progress bar. | Automated KPI dashboards where percentages are rendered with no gaps for a dense visual. | Printable Excel sheets that display task progress without manual chart configuration.
// AI Prompts: How do I bind a range of cells to a chart series and keep the series visible in Aspose.Cells for .NET? | Provide C# code to style a stacked bar chart as a progress bar by setting Overlap to -100 and GapWidth to 0. | Explain how to assign category labels from column A while binding progress values from column B in an Aspose.Cells chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ProgressBarChartExample
{
    // This example creates a workbook, fills column A with task names and column B with numeric progress percentages, adds a stacked bar chart, binds the B2:B10 range to the first series, assigns A2:A10 as category labels, disables series filtering, and sets Overlap = -100 and GapWidth = 0 to render a classic progress‑bar look before saving as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data: categories in column A, progress percentages in column B
            sheet.Cells["A1"].PutValue("Task");
            sheet.Cells["B1"].PutValue("Progress");
            for (int i = 2; i <= 10; i++)
            {
                sheet.Cells[$"A{i}"].PutValue($"Task {i - 1}");
                // Example progress values (0% to 100%)
                sheet.Cells[$"B{i}"].PutValue((i - 1) * 10); // 10,20,...,90
            }

            // Add a stacked bar chart that will act as a progress bar
            int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 12, 0, 30, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Bind the progress percentages to the first (and only) visible series
            // Add the series values from column B
            chart.NSeries.Add("B2:B10", true);
            // Set the category (task names) for the series
            chart.NSeries.CategoryData = "A2:A10";

            // Ensure the series is visible (not filtered)
            chart.NSeries[0].IsFiltered = false;

            // Optional: adjust appearance for a classic progress bar look
            chart.NSeries[0].Overlap = -100; // bars fully overlap
            chart.NSeries[0].GapWidth = 0;   // no gap between bars

            // Save the workbook
            workbook.Save("ProgressBarChart.xlsx", SaveFormat.Xlsx);
        }
    }
}
