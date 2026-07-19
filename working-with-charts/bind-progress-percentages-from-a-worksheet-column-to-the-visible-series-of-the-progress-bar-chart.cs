// Title: C# – Bind Column Percentages to a Progress‑Bar Bar Chart with Aspose.Cells
// Description: Shows how to create a workbook, fill column A with task names and column B with decimal progress values, add a bar chart, link B2:B5 as the series and A2:A5 as categories, enable PlotVisibleCellsOnly, tweak GapWidth and Overlap for a compact bar look, and save the file as ProgressBarChart.xlsx.
// Keywords: Aspose.Cells C# bind chart series | progress bar chart Aspose.Cells | PlotVisibleCellsOnly chart | bar chart gap width overlap | dynamic data range chart Aspose | Excel progress visualization .NET | Aspose.Cells chart customization | bind column to chart series | visible cells only chart | Aspose.Cells bar chart example
// Common Searches: Aspose.Cells bind column to chart series | Create a progress bar chart in C# with Aspose.Cells | Plot only visible rows in Aspose.Cells chart | Set gap width and overlap for bar chart Aspose.Cells | C# example progress bar Excel chart Aspose | Dynamic range for chart series Aspose.Cells
// Developer Intent: Link worksheet column data to a bar‑chart series and render only visible rows to produce a progress‑bar style visualization.
// Use Cases: Build a project‑status dashboard where each task’s completion percentage appears as a horizontal bar. | Generate a printable Excel report that hides completed rows while the chart automatically reflects the visible data. | Create a compact visual indicator for KPI tracking by adjusting bar spacing and overlap.
// AI Prompts: Provide C# code that binds a changing range of progress values in column B to a bar chart series and keeps the chart updated when rows are added or hidden. | Explain how to enable PlotVisibleCellsOnly and customize GapWidth and Overlap to mimic a progress bar using Aspose.Cells. | Show how to apply conditional formatting or color gradients to the bars for a more expressive progress indicator.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Shows how to create a workbook, fill column A with task names and column B with decimal progress values, add a bar chart, link B2:B5 as the series and A2:A5 as categories, enable PlotVisibleCellsOnly, tweak GapWidth and Overlap for a compact bar look, and save the file as ProgressBarChart.xlsx.
class ProgressBarChartExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data: task names and progress percentages
        sheet.Cells["A1"].PutValue("Task");
        sheet.Cells["B1"].PutValue("Progress");
        sheet.Cells["A2"].PutValue("Design");
        sheet.Cells["A3"].PutValue("Development");
        sheet.Cells["A4"].PutValue("Testing");
        sheet.Cells["A5"].PutValue("Deployment");
        sheet.Cells["B2"].PutValue(0.25); // 25%
        sheet.Cells["B3"].PutValue(0.55); // 55%
        sheet.Cells["B4"].PutValue(0.80); // 80%
        sheet.Cells["B5"].PutValue(0.95); // 95%

        // Add a bar chart that will act as a progress bar
        int chartIndex = sheet.Charts.Add(ChartType.Bar, 7, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Bind the progress percentages (column B) to the visible series of the chart
        // Add the series values (vertical orientation) and set category (task names)
        chart.NSeries.Add("B2:B5", true);
        chart.NSeries.CategoryData = "A2:A5";

        // Ensure only visible cells are plotted (useful if rows are hidden later)
        chart.PlotVisibleCellsOnly = true;

        // Optional: adjust visual appearance of the progress bar
        chart.GapWidth = 150;          // space between bars
        chart.NSeries[0].Overlap = -30; // make bars overlap slightly

        // Save the workbook
        workbook.Save("ProgressBarChart.xlsx");
    }
}
