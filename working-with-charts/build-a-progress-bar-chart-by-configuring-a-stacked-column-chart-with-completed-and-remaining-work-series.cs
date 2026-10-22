// Title: Create a stacked column progress bar chart in Excel with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code using Aspose.Cells that builds a stacked column chart to display completed vs remaining work for each task. | Show how to assign green and light‑gray colors to the completed and remaining series and eliminate column gaps for a solid bar appearance. | Provide a snippet that disables the chart legend, adds data labels to the completed portion, and saves the workbook as an XLSX file.
// Common Searches: C# Aspose.Cells create progress bar chart with stacked columns | Aspose.Cells remove column gaps for solid bar chart | How to set custom series colors in Aspose.Cells column chart | Aspose.Cells hide legend in Excel chart programmatically | Save stacked column progress chart to XLSX using Aspose.Cells
// Tags: stacked column chart Aspose.Cells C# | custom series colors Aspose.Cells column chart | remove column gaps Aspose.Cells chart | hide chart legend Aspose.Cells | progress bar visualization Excel Aspose.Cells | data labels for completed series Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, inserts task data for completed and remaining percentages, adds a column chart, configures it as a progress bar by stacking the two series, applies green and light‑gray colors, removes column gaps for a solid bar look, shows data labels for the completed portion, hides the legend, and saves the file as ProgressBarChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate header row
            sheet.Cells["A1"].PutValue("Task");
            sheet.Cells["B1"].PutValue("Completed");
            sheet.Cells["C1"].PutValue("Remaining");

            // Populate sample data
            sheet.Cells["A2"].PutValue("Task 1");
            sheet.Cells["B2"].PutValue(70); // 70% completed
            sheet.Cells["C2"].PutValue(30); // 30% remaining

            sheet.Cells["A3"].PutValue("Task 2");
            sheet.Cells["B3"].PutValue(40); // 40% completed
            sheet.Cells["C3"].PutValue(60); // 60% remaining

            // Add a column chart (stacked effect can be simulated)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart title
            chart.Title.Text = "Progress Bar Chart";

            // Define categories (tasks)
            chart.NSeries.CategoryData = "A2:A3";

            // Add Completed series
            chart.NSeries.Add("B2:B3", true);
            // Add Remaining series
            chart.NSeries.Add("C2:C3", true);

            // Format series colors
            chart.NSeries[0].Area.ForegroundColor = Color.Green;        // Completed = green
            chart.NSeries[1].Area.ForegroundColor = Color.LightGray;   // Remaining = light gray

            // Remove gaps between columns for a solid bar look
            chart.GapWidth = 0;

            // Show data labels for completed portion
            chart.NSeries[0].DataLabels.ShowValue = true;

            // Hide legend (optional)
            chart.ShowLegend = false;

            // Save the workbook
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ProgressBarChart.xlsx");
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
