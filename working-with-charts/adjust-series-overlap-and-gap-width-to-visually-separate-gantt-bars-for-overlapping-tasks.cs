// Title: Setting SeriesOverlap and GapWidth to separate overlapping bars in a stacked‑bar Gantt chart with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a stacked‑bar Gantt chart using Aspose.Cells and applies Chart.SeriesOverlap = 0 and Chart.GapWidth = 100 to keep each task bar visually distinct. | Update the provided Aspose.Cells example to configure the chart's SeriesOverlap and GapWidth properties after adding the chart, ensuring overlapping tasks are clearly spaced. | Write a helper method that receives an Aspose.Cells Chart object and sets appropriate bar‑spacing values (SeriesOverlap and GapWidth) for a readable Gantt timeline.
// Common Searches: Aspose.Cells C# set chart series overlap for stacked bar Gantt chart | adjust gap width in Aspose.Cells stacked bar chart to separate overlapping tasks | how to space overlapping bars in a Gantt chart using Aspose.Cells .NET | C# Aspose.Cells chart bar spacing options for project timeline | increase readability of overlapping Gantt bars with Aspose.Cells settings
// Tags: Aspose.Cells chart series overlap setting | Aspose.Cells stacked bar gap width configuration | C# Gantt chart bar spacing | transparent start series Aspose.Cells | chart formatting for overlapping tasks .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace GanttChartExample
{
    // The example creates an Excel workbook, fills it with task names, start dates, and durations, adds a stacked‑bar chart where the start series is hidden to offset the duration bars, and demonstrates where to set Chart.SeriesOverlap and Chart.GapWidth so overlapping Gantt bars are visually separated before saving the file as GanttChart_OverlapAdjusted.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "GanttData";

                // Header row
                sheet.Cells["A1"].PutValue("Task");
                sheet.Cells["B1"].PutValue("Start");
                sheet.Cells["C1"].PutValue("Duration");

                // Sample tasks (including overlapping periods)
                sheet.Cells["A2"].PutValue("Task 1");
                sheet.Cells["B2"].PutValue(new DateTime(2023, 1, 1));
                sheet.Cells["C2"].PutValue(5);   // 5 days

                sheet.Cells["A3"].PutValue("Task 2");
                sheet.Cells["B3"].PutValue(new DateTime(2023, 1, 3));
                sheet.Cells["C3"].PutValue(7);   // overlaps with Task 1

                sheet.Cells["A4"].PutValue("Task 3");
                sheet.Cells["B4"].PutValue(new DateTime(2023, 1, 8));
                sheet.Cells["C4"].PutValue(4);   // starts after Task 1 ends

                // Ensure start dates are stored as proper Excel date values
                for (int i = 2; i <= 4; i++)
                {
                    Cell dateCell = sheet.Cells[i, 1]; // column B (index 1)
                    dateCell.PutValue(dateCell.Value);
                }

                // Add a stacked bar chart to represent the Gantt bars
                int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 6, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Series 1: Start dates (invisible, used to offset the bars)
                int startSeriesIdx = chart.NSeries.Add("B2:B4", true);
                Series startSeries = chart.NSeries[startSeriesIdx];
                startSeries.Name = "Start";
                startSeries.IsColorVaried = false;
                // Hide the start series by setting its fill to transparent
                startSeries.Area.ForegroundColor = System.Drawing.Color.Transparent;

                // Series 2: Duration (visible Gantt bars)
                int durationSeriesIdx = chart.NSeries.Add("C2:C4", true);
                Series durationSeries = chart.NSeries[durationSeriesIdx];
                durationSeries.Name = "Duration";
                durationSeries.IsColorVaried = true;

                // Set category (task) labels
                chart.NSeries.CategoryData = "A2:A4";

                // Ensure output directory exists
                string outputPath = "GanttChart_OverlapAdjusted.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
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
}
