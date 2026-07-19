// Title: C# Aspose.Cells Example: Adjust Series Overlap & Gap Width in a Stacked Bar Gantt Chart
// Description: Shows how to build an Excel workbook with task data, insert a stacked bar chart for a Gantt view, and apply NSeries.Overlap (negative) and NSeries.GapWidth settings to visually separate overlapping task bars before saving the file.
// Keywords: Aspose.Cells | C# | .NET | stacked bar chart | series overlap | gap width | Gantt chart | overlapping tasks | chart spacing | Excel automation | negative overlap | chart gap width | GitHub example
// Common Searches: Aspose.Cells set series overlap C# | increase gap width stacked bar chart Aspose.Cells | separate overlapping Gantt bars Aspose.Cells .NET | negative overlap value chart series Aspose | adjust bar spacing in Excel chart using Aspose.Cells
// Developer Intent: Configure NSeries.Overlap and NSeries.GapWidth to create clear spacing between overlapping bars in a stacked bar Gantt chart.
// Use Cases: Generate a project schedule where tasks that start at similar times are displayed with a visible gap. | Create printable Gantt charts with distinct bars for overlapping activities. | Automate Excel reports that require precise visual separation of task durations.
// AI Prompts: Provide C# code that creates a stacked bar Gantt chart with Aspose.Cells and sets a negative series overlap and increased gap width. | Explain how NSeries.Overlap and NSeries.GapWidth affect the appearance of overlapping bars in an Aspose.Cells chart. | Show a step‑by‑step example of adjusting bar spacing for a Gantt chart using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace GanttChartExample
{
    // Shows how to build an Excel workbook with task data, insert a stacked bar chart for a Gantt view, and apply NSeries.Overlap (negative) and NSeries.GapWidth settings to visually separate overlapping task bars before saving the file.
    class GanttChartDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data: Task name, Start (offset), Duration (length)
                sheet.Cells["A1"].PutValue("Task");
                sheet.Cells["B1"].PutValue("Start");
                sheet.Cells["C1"].PutValue("Duration");

                sheet.Cells["A2"].PutValue("Task 1");
                sheet.Cells["B2"].PutValue(1);
                sheet.Cells["C2"].PutValue(3);

                sheet.Cells["A3"].PutValue("Task 2");
                sheet.Cells["B3"].PutValue(2);
                sheet.Cells["C3"].PutValue(4);

                sheet.Cells["A4"].PutValue("Task 3");
                sheet.Cells["B4"].PutValue(3);
                sheet.Cells["C4"].PutValue(2);

                // Add a stacked bar chart (commonly used for Gantt charts)
                int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // First series: Start offsets (will be hidden later if needed)
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries[0].Name = "Start";

                // Second series: Durations (visible bars)
                chart.NSeries.Add("C2:C4", false);
                chart.NSeries[1].Name = "Duration";

                // Set category (task) labels
                chart.NSeries.CategoryData = "A2:A4";

                // Adjust series overlap to separate overlapping bars (negative value)
                chart.NSeries[0].Overlap = -50;
                chart.NSeries[1].Overlap = -50;

                // Increase gap width to add spacing between bar clusters
                chart.NSeries[0].GapWidth = 200;
                chart.NSeries[1].GapWidth = 200;

                // Define output file path
                string outputPath = "GanttChartOverlapGap.xlsx";

                // Ensure the directory exists (use current directory if none)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook with the configured chart
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the Gantt chart:");
                Console.WriteLine(ex.Message);
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            GanttChartDemo.Run();
        }
    }
}
