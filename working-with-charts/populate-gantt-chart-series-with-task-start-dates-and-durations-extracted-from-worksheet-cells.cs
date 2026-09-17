// Title: Create a Gantt chart in Excel using Aspose.Cells for .NET by loading task names, start dates, and durations from worksheet cells
// AI Prompts: Generate C# code with Aspose.Cells that reads task names, start dates, and duration values from a worksheet range and builds a stacked‑bar Gantt chart on a new sheet. | Adjust the chart so the start‑date series is hidden (transparent) and the duration series uses a light‑blue fill, while assigning the task names to the vertical axis labels. | Add data labels to the duration series that display each task's duration on the Gantt bars and save the workbook to a specified output file.
// Common Searches: aspnet create gantt chart from excel data using Aspose.Cells stacked bar | how to bind start date and duration columns to a Gantt chart in Aspose.Cells C# | make start date series invisible in Aspose.Cells stacked bar chart | set task names as category axis labels in Aspose.Cells chart | save generated Gantt chart to a new worksheet with Aspose.Cells .NET
// Tags: Aspose.Cells stacked bar Gantt chart creation | populate chart series from Excel range C# | hide start date series Aspose.Cells | task names as category axis labels Aspose.Cells | save workbook with generated chart .NET

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads task names, start dates, and durations from the first worksheet of an input Excel file, creates a stacked‑bar Gantt chart on a new sheet, makes the start‑date series transparent, colors the duration series light blue, sets task names as vertical axis labels, and saves the workbook with the chart.
class GanttChartGenerator
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the source workbook that contains task data
            Workbook workbook = new Workbook(inputPath);

            // Assume task data is on the first worksheet
            Worksheet dataSheet = workbook.Worksheets[0];

            // Add a new worksheet for the Gantt chart
            int chartSheetIndex = workbook.Worksheets.Add();
            Worksheet chartSheet = workbook.Worksheets[chartSheetIndex];
            chartSheet.Name = "Gantt";

            // Define the range of data (A: Task Name, B: Start Date, C: Duration)
            int firstDataRow = 1; // zero‑based index, row 2 in Excel (skip header)
            int lastDataRow = dataSheet.Cells.MaxDataRow; // last row with data

            string taskNamesRange = $"'{dataSheet.Name}'!A{firstDataRow + 1}:A{lastDataRow + 1}";
            string startDatesRange = $"'{dataSheet.Name}'!B{firstDataRow + 1}:B{lastDataRow + 1}";
            string durationsRange = $"'{dataSheet.Name}'!C{firstDataRow + 1}:C{lastDataRow + 1}";

            // Add a stacked bar chart (used for Gantt representation)
            int chartIndex = chartSheet.Charts.Add(ChartType.BarStacked, 1, 1, 20, 10);
            Chart ganttChart = chartSheet.Charts[chartIndex];

            // First series: Start Dates (invisible, used to offset the bars)
            int startSeriesIdx = ganttChart.NSeries.Add(startDatesRange, true);
            Series startSeries = ganttChart.NSeries[startSeriesIdx];
            startSeries.Area.ForegroundColor = Color.Transparent;
            startSeries.Border.IsVisible = false;

            // Second series: Durations (visible bars)
            int durationSeriesIdx = ganttChart.NSeries.Add(durationsRange, true);
            Series durationSeries = ganttChart.NSeries[durationSeriesIdx];
            durationSeries.Area.ForegroundColor = Color.LightBlue;

            // Set the category (vertical) axis labels to task names
            ganttChart.NSeries.CategoryData = taskNamesRange;

            // Configure legend position
            if (ganttChart.Legend != null)
            {
                ganttChart.Legend.Position = LegendPositionType.Bottom;
            }

            // Hide the start series from the legend
            ganttChart.NSeries[0].Name = string.Empty;

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the new Gantt chart
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Gantt chart generated and saved to \"{outputPath}\".");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
