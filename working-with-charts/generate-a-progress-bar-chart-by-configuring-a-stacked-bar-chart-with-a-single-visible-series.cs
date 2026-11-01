// Title: Create a progress bar visualization using a stacked bar chart with a hidden series in Aspose.Cells for .NET (C#)
// AI Prompts: Generate a stacked bar chart in Aspose.Cells, make the second series transparent, and export the workbook as an .xlsx file using C#. | Insert task, progress, and remaining data into a worksheet, conceal the legend and value axis, and set a custom chart title for a progress‑bar effect. | Adjust the chart's position and size, then save the workbook containing the progress bar chart to a specified directory in C#.
// Common Searches: Aspose.Cells C# create progress bar chart using stacked bar and hide remaining series | how to make a chart series transparent in Aspose.Cells for .NET | remove legend and value axis from Aspose.Cells stacked bar chart C# | save workbook with progress bar visualization Aspose.Cells .NET | configure chart dimensions for progress bar in Aspose.Cells C#
// Tags: stacked bar chart progress visualization Aspose.Cells | transparent series Aspose.Cells C# | remove chart legend Aspose.Cells | disable value axis Aspose.Cells | save workbook as xlsx Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds task, progress, and remaining data, inserts a stacked bar chart, makes the remaining series transparent to simulate a progress bar, hides the legend and value axis, customizes the title, and saves the file as ProgressBarChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data for progress and remaining portions
            sheet.Cells["A1"].PutValue("Task");
            sheet.Cells["B1"].PutValue("Progress");
            sheet.Cells["C1"].PutValue("Remaining");

            sheet.Cells["A2"].PutValue("Project A");
            sheet.Cells["B2"].PutValue(70); // 70% completed
            sheet.Cells["C2"].PutValue(30); // 30% remaining

            sheet.Cells["A3"].PutValue("Project B");
            sheet.Cells["B3"].PutValue(45); // 45% completed
            sheet.Cells["C3"].PutValue(55); // 55% remaining

            // Add a stacked bar chart (acts as a progress bar)
            int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart title
            chart.Title.Text = "Progress Bar Chart";

            // Add the visible series (Progress)
            int progressSeriesIdx = chart.NSeries.Add("B2:B3", true);

            // Add the hidden series (Remaining)
            int remainingSeriesIdx = chart.NSeries.Add("C2:C3", true);

            // Hide the remaining series to simulate a progress bar
            chart.NSeries[remainingSeriesIdx].Area.ForegroundColor = Color.Transparent;

            // Hide the entire legend for a cleaner look
            chart.ShowLegend = false;

            // Optional: hide the value axis for a cleaner look
            chart.ValueAxis.IsVisible = false;

            // Determine output path and ensure directory exists
            string outputPath = "ProgressBarChart.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the chart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
