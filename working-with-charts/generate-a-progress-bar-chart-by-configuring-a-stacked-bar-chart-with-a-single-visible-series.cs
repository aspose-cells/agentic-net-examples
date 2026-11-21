// Title: How to generate a progress bar chart using a stacked bar chart with a single visible series in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# example for creating a progress bar using a stacked bar chart | set transparent background series in Aspose.Cells stacked bar chart | remove gaps between bars in Aspose.Cells chart for solid progress bar | change series fill color to green in Aspose.Cells progress bar chart | save progress bar chart as Excel file with Aspose.Cells .NET
// Tags: stacked bar chart transparent background series Aspose.Cells | configure chart gap width zero Aspose.Cells | set series fill color green Aspose.Cells chart | export progress bar chart to XLSX Aspose.Cells | single visible series stacked bar Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, adds task data with remaining and completed values, inserts a stacked bar chart, makes the background series transparent, colors the progress series green, removes bar gaps for a solid look, and saves the result as ProgressBarChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet's cells
            var workbook = new Workbook();
            var cells = workbook.Worksheets[0].Cells;

            // -------------------------------------------------
            // Prepare data for the progress bar (stacked bar)
            // -------------------------------------------------
            // Header row
            cells["A1"].PutValue("Task");
            cells["B1"].PutValue("Remaining");   // will be hidden
            cells["C1"].PutValue("Completed");   // visible part

            // Data row
            cells["A2"].PutValue("Project A");
            double total = 100;          // total length of the bar
            double completed = 70;       // progress value
            cells["B2"].PutValue(total - completed); // invisible remainder
            cells["C2"].PutValue(completed);          // visible progress

            // -------------------------------------------------
            // Create a stacked bar chart
            // -------------------------------------------------
            var worksheet = workbook.Worksheets[0];
            // Add chart at position (row 5, column 0) to (row 20, column 10)
            int chartIdx = worksheet.Charts.Add(ChartType.BarStacked, 5, 0, 20, 10);
            var chart = worksheet.Charts[chartIdx];

            // Optional: set a title
            chart.Title.Text = "Progress Bar";

            // -------------------------------------------------
            // Series 1 – Background (invisible)
            // -------------------------------------------------
            int bgSeriesIdx = chart.NSeries.Add("B2", true);
            var bgSeries = chart.NSeries[bgSeriesIdx];
            // Make the background series transparent
            bgSeries.Area.ForegroundColor = Color.Transparent;
            bgSeries.Border.IsVisible = false;

            // -------------------------------------------------
            // Series 2 – Progress (visible)
            // -------------------------------------------------
            int progSeriesIdx = chart.NSeries.Add("C2", true);
            var progSeries = chart.NSeries[progSeriesIdx];
            progSeries.Area.ForegroundColor = Color.Green; // color of the progress bar
            progSeries.Border.IsVisible = false;

            // Remove gaps between bars for a solid look
            chart.GapWidth = 0;

            // -------------------------------------------------
            // Save the workbook
            // -------------------------------------------------
            string outputPath = "ProgressBarChart.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? string.Empty;

            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
