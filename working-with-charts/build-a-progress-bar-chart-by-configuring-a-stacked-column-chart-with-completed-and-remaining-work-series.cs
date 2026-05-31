using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

class ProgressBarChart
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // -------------------------------------------------
        // Populate sample data: Task name, Completed %, Remaining %
        // -------------------------------------------------
        sheet.Cells["A1"].PutValue("Task");
        sheet.Cells["B1"].PutValue("Completed");
        sheet.Cells["C1"].PutValue("Remaining");

        string[] tasks = { "Task 1", "Task 2", "Task 3" };
        double[] completed = { 70, 45, 90 };

        for (int i = 0; i < tasks.Length; i++)
        {
            int row = i + 2; // Data starts from row 2
            sheet.Cells[row, 0].PutValue(tasks[i]);               // Column A
            sheet.Cells[row, 1].PutValue(completed[i]);           // Column B
            sheet.Cells[row, 2].PutValue(100 - completed[i]);    // Column C (Remaining)
        }

        // -------------------------------------------------
        // Add a stacked column chart (acts as a progress bar)
        // -------------------------------------------------
        int chartIndex = sheet.Charts.Add(ChartType.ColumnStacked, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Add the two series: Completed and Remaining
        chart.NSeries.Add("B2:B4", true); // Completed series
        chart.NSeries.Add("C2:C4", true); // Remaining series

        // Set the category (X‑axis) labels
        chart.NSeries.CategoryData = "A2:A4";

        // -------------------------------------------------
        // Visual tweaks
        // -------------------------------------------------
        // Remove gaps between the stacked columns
        chart.GapWidth = 0; // Chart.GapWidth property

        // Optional: set overlap to fully overlay (demonstrates Series.Overlap property)
        chart.NSeries[0].Overlap = -100; // Series.Overlap property

        // Assign colors: Completed = green, Remaining = light gray
        chart.NSeries[0].Area.ForegroundColor = Color.Green;
        chart.NSeries[1].Area.ForegroundColor = Color.LightGray;

        // Add a chart title
        chart.Title.Text = "Progress Bar Chart";

        // -------------------------------------------------
        // Save the workbook
        // -------------------------------------------------
        workbook.Save("ProgressBarChart.xlsx");
    }
}