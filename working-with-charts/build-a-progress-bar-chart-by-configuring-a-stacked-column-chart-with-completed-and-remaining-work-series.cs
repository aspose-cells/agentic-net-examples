// Title: Build a Progress‑Bar Style Stacked Column Chart with Aspose.Cells for .NET (C#)
// Description: This example shows how to create a new workbook, populate task, completed, and remaining data, add a stacked column chart, remove column gaps, set full overlap, apply green and light‑gray colors, hide the legend, add a title, and save the file as ProgressBarChart.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# progress bar chart | stacked column chart Aspose.Cells | remove gaps Aspose.Cells chart | chart series overlap Aspose.Cells | custom series colors Aspose.Cells | Excel progress bar .NET | Aspose.Cells chart formatting | C# generate progress bar Excel
// Common Searches: Aspose.Cells create progress bar chart | stacked column chart without gaps C# | set overlap for Aspose.Cells column series | change chart series colors Aspose.Cells | C# Aspose.Cells progress bar example
// Developer Intent: Generate a stacked column chart that looks like a progress bar, displaying completed versus remaining work for each task in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Show task completion percentages in a project status report. | Add visual progress indicators to a dashboard sheet with multiple processes. | Export a concise progress summary to XLSX for stakeholder distribution.
// AI Prompts: Write C# code with Aspose.Cells to create a stacked column chart that mimics a progress bar, using custom colors and no gaps. | Explain how to configure GapWidth and Overlap properties in Aspose.Cells to produce seamless bars. | Show how to add data labels and format them for a progress‑bar style chart in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This example shows how to create a new workbook, populate task, completed, and remaining data, add a stacked column chart, remove column gaps, set full overlap, apply green and light‑gray colors, hide the legend, add a title, and save the file as ProgressBarChart.xlsx using Aspose.Cells for .NET.
class ProgressBarChartDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // -------------------------------------------------
        // Populate sample data:
        // Column A – Task names (categories)
        // Column B – Completed work
        // Column C – Remaining work
        // -------------------------------------------------
        sheet.Cells["A1"].PutValue("Task");
        sheet.Cells["B1"].PutValue("Completed");
        sheet.Cells["C1"].PutValue("Remaining");

        string[] tasks = { "Design", "Development", "Testing", "Deployment" };
        double[] completed = { 70, 40, 20, 10 };
        double[] remaining = { 30, 60, 80, 90 };

        for (int i = 0; i < tasks.Length; i++)
        {
            int row = i + 2; // Data starts from row 2
            sheet.Cells[$"A{row}"].PutValue(tasks[i]);
            sheet.Cells[$"B{row}"].PutValue(completed[i]);
            sheet.Cells[$"C{row}"].PutValue(remaining[i]);
        }

        // -------------------------------------------------
        // Add a stacked column chart (acts as a progress bar)
        // -------------------------------------------------
        int chartIndex = sheet.Charts.Add(ChartType.ColumnStacked, 6, 0, 25, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the category (X‑axis) data
        chart.NSeries.CategoryData = "A2:A5";

        // Add the "Completed" series
        chart.NSeries.Add("B2:B5", true);
        // Add the "Remaining" series
        chart.NSeries.Add("C2:C5", true);

        // -------------------------------------------------
        // Visual tweaks to make it look like a progress bar
        // -------------------------------------------------
        // Remove gaps between columns
        chart.GapWidth = 0;

        // Optional: increase overlap so the two series appear as a single bar
        // (only effective for 2‑D column charts)
        chart.NSeries[0].Overlap = 100;
        chart.NSeries[1].Overlap = 100;

        // Set colors: Completed – green, Remaining – light gray
        chart.NSeries[0].Area.ForegroundColor = Color.Green;
        chart.NSeries[1].Area.ForegroundColor = Color.LightGray;

        // Hide the legend (not needed for a simple progress bar)
        chart.ShowLegend = false;

        // Add a title
        chart.Title.Text = "Project Progress";

        // -------------------------------------------------
        // Save the workbook
        // -------------------------------------------------
        workbook.Save("ProgressBarChart.xlsx", SaveFormat.Xlsx);
    }
}
