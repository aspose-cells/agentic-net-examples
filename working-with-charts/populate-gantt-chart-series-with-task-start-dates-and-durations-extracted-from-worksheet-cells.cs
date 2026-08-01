// Title: Generate a Gantt‑style stacked bar chart with Aspose.Cells C# from worksheet data
// Description: Creates a new workbook, writes task names, start dates and durations into cells, adds a stacked bar chart, binds the duration range as series values, uses the start‑date range as XValues, sets task names as category labels, and saves the file as an Excel Gantt‑style timeline.
// Keywords: Aspose.Cells | C# | Gantt chart | stacked bar chart | XValues | CategoryData | project timeline | Excel automation | populate chart series | date axis
// Common Searches: Aspose.Cells C# Gantt chart example | How to bind start dates to X axis in Aspose.Cells chart | Create stacked bar chart with dates using Aspose.Cells | Set category labels for chart series Aspose.Cells | Generate project timeline Excel with Aspose.Cells
// Developer Intent: Build an Excel workbook that displays a project schedule as a Gantt‑style stacked bar chart by reading task data from worksheet cells.
// Use Cases: Automatically produce visual project schedules for multiple initiatives directly from Excel data. | Export a printable timeline report for stakeholders without manual chart configuration. | Integrate Gantt chart generation into CI/CD pipelines that deliver project status updates in Excel format.
// AI Prompts: Add a second series to show completed work and color it differently in the Gantt chart. | Format the X‑axis to display calendar dates instead of serial numbers in Aspose.Cells. | Create conditional formatting that changes bar colors based on task status (e.g., completed, in‑progress). | Generate a dynamic Gantt chart that expands when new rows are added to the task table.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace GanttChartDemo
{
    // Creates a new workbook, writes task names, start dates and durations into cells, adds a stacked bar chart, binds the duration range as series values, uses the start‑date range as XValues, sets task names as category labels, and saves the file as an Excel Gantt‑style timeline.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Header row
                sheet.Cells["A1"].PutValue("Task");
                sheet.Cells["B1"].PutValue("Start Date");
                sheet.Cells["C1"].PutValue("Duration");

                // Sample tasks
                sheet.Cells["A2"].PutValue("Design");
                sheet.Cells["B2"].PutValue(new DateTime(2023, 1, 1));
                sheet.Cells["C2"].PutValue(10);

                sheet.Cells["A3"].PutValue("Development");
                sheet.Cells["B3"].PutValue(new DateTime(2023, 1, 12));
                sheet.Cells["C3"].PutValue(20);

                sheet.Cells["A4"].PutValue("Testing");
                sheet.Cells["B4"].PutValue(new DateTime(2023, 2, 5));
                sheet.Cells["C4"].PutValue(8);

                sheet.Cells["A5"].PutValue("Deployment");
                sheet.Cells["B5"].PutValue(new DateTime(2023, 2, 15));
                sheet.Cells["C5"].PutValue(4);

                // Add a stacked bar chart (used to simulate a Gantt chart)
                int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Add series for durations (Y values)
                int seriesIdx = chart.NSeries.Add("C2:C5", true);
                Series series = chart.NSeries[seriesIdx];

                // Set start dates as X values
                series.XValues = "B2:B5";

                // Set task names as category (Y‑axis) labels
                chart.NSeries.CategoryData = "A2:A5";

                // Optional series name
                series.Name = "Project Timeline";

                // Save the workbook
                string outputPath = "GanttChartDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
