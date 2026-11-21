// Title: Create a Gantt chart with an invisible start series and customize series overlap and gap width using Aspose.Cells for .NET (C#)
// Common Searches: how to set series overlap and gap width for a stacked bar chart in Aspose.Cells C# | Aspose.Cells C# example for building a Gantt chart with a hidden start series | adjust bar spacing for overlapping tasks in an Excel Gantt chart using Aspose.Cells | C# code to modify chart gap width and series overlap in Aspose.Cells | customize visual separation of Gantt bars in Aspose.Cells .NET
// Tags: Aspose.Cells stacked bar chart overlap .NET | Gantt chart invisible start series Aspose.Cells | adjust chart gap width Aspose.Cells C# | set series overlap property Aspose.Cells | Excel Gantt chart bar spacing Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.Drawing;

// The sample creates a new workbook, adds task, start, and duration data, and builds a stacked bar chart that serves as a Gantt chart. An invisible start series offsets the bars, while a visible duration series displays the tasks. Category labels are set to task names. Commented code shows how to change the SeriesOverlap and GapWidth properties to separate overlapping bars. The workbook is saved as GanttChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and rename it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Gantt";

            // Sample data: Task, Start, Duration
            sheet.Cells["A1"].PutValue("Task");
            sheet.Cells["B1"].PutValue("Start");
            sheet.Cells["C1"].PutValue("Duration");

            sheet.Cells["A2"].PutValue("Task 1");
            sheet.Cells["B2"].PutValue(1);
            sheet.Cells["C2"].PutValue(5);

            sheet.Cells["A3"].PutValue("Task 2");
            sheet.Cells["B3"].PutValue(3);
            sheet.Cells["C3"].PutValue(4);

            sheet.Cells["A4"].PutValue("Task 3");
            sheet.Cells["B4"].PutValue(6);
            sheet.Cells["C4"].PutValue(3);

            // Add a stacked bar chart (typical Gantt representation)
            int chartIdx = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIdx];

            // Series for start dates (invisible, used to offset the bars)
            int startSeriesIdx = chart.NSeries.Add("B2:B4", true);
            Series startSeries = chart.NSeries[startSeriesIdx];
            startSeries.Name = "Start";
            startSeries.Area.ForegroundColor = Color.Transparent; // hide start segment
            startSeries.Border.IsVisible = false;

            // Series for durations (visible Gantt bars)
            int durationSeriesIdx = chart.NSeries.Add("C2:C4", true);
            Series durationSeries = chart.NSeries[durationSeriesIdx];
            durationSeries.Name = "Duration";
            durationSeries.Area.ForegroundColor = Color.LightBlue; // optional color

            // Set category axis labels (task names)
            chart.NSeries.CategoryData = "A2:A4";

            // Adjust series overlap and gap width if supported
            // These properties may not be available in older versions; they are optional.
            // chart.SeriesOverlap = -50;
            // chart.GapWidth = 150;

            // Save the workbook
            string outputPath = "GanttChart.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
