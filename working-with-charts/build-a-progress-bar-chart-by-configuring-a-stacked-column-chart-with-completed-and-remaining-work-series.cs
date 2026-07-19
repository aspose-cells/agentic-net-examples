// Title: C# – Create a Progress‑Bar Stacked Column Chart with Aspose.Cells
// Description: This example shows how to generate an Excel workbook in .NET, add task names with completed and remaining percentages, and render them as a progress‑bar using a stacked column chart. It configures series overlap, removes column gaps, sets the category axis, adds a chart title, and saves the file as ProgressBarChart.xlsx.
// Keywords: Aspose.Cells stacked column chart | progress bar chart C# | series overlap Aspose.Cells | gap width zero Excel chart | visualize task completion Aspose.Cells | .NET Excel progress bar | Aspose.Cells chart customization
// Common Searches: how to make a progress bar chart with Aspose.Cells C# | stacked column chart overlap and gap width Aspose.Cells | Aspose.Cells create progress bar for project tasks | C# code for Excel progress bar using stacked columns
// Developer Intent: Produce an Excel file that displays each task’s completion percentage as a compact progress‑bar using a stacked column chart.
// Use Cases: Project status reports that show task progress at a glance | Dashboard sheets with resource‑utilization bars for executives | Automated export of daily metrics where each row appears as a progress bar
// AI Prompts: Generate C# code with Aspose.Cells to build a stacked column chart that looks like a progress bar, setting series overlap to 100 % and gap width to 0. | Write a method that accepts a list of tasks (completed and remaining values) and returns a workbook containing a formatted progress‑bar chart. | Explain how to assign custom foreground colors to the completed and remaining series in an Aspose.Cells stacked column chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ProgressBarChartDemo
{
    // This example shows how to generate an Excel workbook in .NET, add task names with completed and remaining percentages, and render them as a progress‑bar using a stacked column chart. It configures series overlap, removes column gaps, sets the category axis, adds a chart title, and saves the file as ProgressBarChart.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Prepare data for a progress bar (stacked column)
            // Column A : Task names (categories)
            // Column B : Completed work
            // Column C : Remaining work
            // -------------------------------------------------
            sheet.Cells["A1"].PutValue("Task");
            sheet.Cells["B1"].PutValue("Completed");
            sheet.Cells["C1"].PutValue("Remaining");

            // Example data for three tasks
            sheet.Cells["A2"].PutValue("Task 1");
            sheet.Cells["B2"].PutValue(70);   // 70% completed
            sheet.Cells["C2"].PutValue(30);   // 30% remaining

            sheet.Cells["A3"].PutValue("Task 2");
            sheet.Cells["B3"].PutValue(45);
            sheet.Cells["C3"].PutValue(55);

            sheet.Cells["A4"].PutValue("Task 3");
            sheet.Cells["B4"].PutValue(90);
            sheet.Cells["C4"].PutValue(10);

            // -------------------------------------------------
            // Add a stacked column chart
            // -------------------------------------------------
            int chartIndex = sheet.Charts.Add(ChartType.ColumnStacked, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the first series (Completed) and second series (Remaining)
            // The first series uses the range B2:B4, the second uses C2:C4
            chart.NSeries.Add("B2:B4", true);          // Completed series
            chart.NSeries.Add("C2:C4", true);          // Remaining series

            // Set category (X‑axis) data
            chart.NSeries.CategoryData = "A2:A4";

            // -------------------------------------------------
            // Visual adjustments to make it look like a progress bar
            // -------------------------------------------------
            // Overlap the two series completely
            chart.NSeries[0].Overlap = 100;   // Fully overlap the stacked columns
            // Remove gaps between columns
            chart.GapWidth = 0;               // No space between bars

            // Optional: give distinct colors (default colors are fine)
            // chart.NSeries[0].Area.ForegroundColor = System.Drawing.Color.Green;
            // chart.NSeries[1].Area.ForegroundColor = System.Drawing.Color.LightGray;

            // Add a title
            chart.Title.Text = "Project Progress";

            // -------------------------------------------------
            // Save the workbook
            // -------------------------------------------------
            workbook.Save("ProgressBarChart.xlsx");
        }
    }
}
