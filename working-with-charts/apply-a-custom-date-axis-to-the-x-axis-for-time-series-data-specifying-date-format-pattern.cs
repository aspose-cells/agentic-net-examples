// Title: Create a line chart with a custom-formatted date X‑axis for time‑series data using Aspose.Cells in C#
// Common Searches: Aspose.Cells C# set custom date format for chart X axis | How to format dates on the X axis of a line chart with Aspose.Cells | Create time‑series line chart with specific date pattern using Aspose.Cells | Example of applying a date number format to a chart axis in Aspose.Cells C# | Plot dates from worksheet cells as X values with custom format in Aspose.Cells
// Tags: date axis formatting Aspose.Cells chart | line chart time series Aspose.Cells | set X axis number format C# Aspose.Cells | populate worksheet with sequential dates Aspose.Cells | custom chart axis pattern Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// The sample creates a new workbook, fills column A with ten consecutive dates starting on 2023‑01‑01 and column B with numeric values, adds a line chart, assigns the date range A1:A10 as the X‑values, applies a custom date number format to the X‑axis, and saves the file as DateAxisChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample time‑series data (dates in column A, values in column B)
            DateTime startDate = new DateTime(2023, 1, 1);
            for (int i = 0; i < 10; i++)
            {
                cells[i, 0].PutValue(startDate.AddDays(i));   // Date
                cells[i, 1].PutValue(i * 10 + 5);            // Value
            }

            // Add a line chart to the worksheet.
            // Parameters: chart type, upper‑left row, upper‑left column, lower‑right row, lower‑right column
            int chartIdx = sheet.Charts.Add(ChartType.Line, 2, 3, 20, 12);
            Chart chart = sheet.Charts[chartIdx];

            // Add series: Y‑values from B1:B10, X‑values (dates) from A1:A10
            int seriesIdx = chart.NSeries.Add("B1:B10", true);
            chart.NSeries[seriesIdx].XValues = "A1:A10";

            // Save the workbook with the chart
            workbook.Save("DateAxisChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
