using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet sheet = workbook.Worksheets[0];

        // Add a stacked bar chart (used for Gantt charts)
        // Parameters: chart type, top row, left column, bottom row, right column
        int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 1, 25, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the chart.
        // Assume column A contains task names, B contains start dates, C contains durations.
        chart.SetChartDataRange("A2:C5", true);

        // Ensure the chart type is set to stacked bar
        chart.Type = ChartType.BarStacked;

        // Configure the chart to look like a Gantt chart
        // Overlap of -100% makes the start and duration series start at the same position
        chart.NSeries[0].Overlap = -100; // Start series
        chart.NSeries[1].Overlap = -100; // Duration series

        // Reduce the gap between bars to zero
        chart.GapWidth = 0;

        // Save the modified workbook
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}