// Title: Create a timeline‑style line chart with percentage data labels, customize label font size, and save it as a PNG image using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that builds a line chart from a date column and a numeric column, enables both value and percentage data labels, formats the labels as percentages, sets the label font size to 12 points, and exports the chart to a PNG file. | Modify the chart to display data labels with a custom number format like '0.0%' and change the label font color to red while keeping the PNG export unchanged. | Add a second series to the timeline chart, turn on data labels for both series, and generate separate PNG files for each series using Aspose.Cells.
// Common Searches: how to show percentage data labels on a line chart with Aspose.Cells C# | set font size for chart data labels in Aspose.Cells .NET | export Aspose.Cells chart as PNG image from a workbook | create timeline chart using line series in Aspose.Cells | format chart data labels as percentages in Aspose.Cells
// Tags: line chart label percentage formatting Aspose.Cells | save chart to PNG file Aspose.Cells .NET | chart data label font size adjustment C# | timeline chart built with line series Aspose.Cells | custom number format for chart labels Aspose.Cells | multiple series handling in Aspose.Cells chart

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, fills it with dates and values, adds a line chart as a timeline substitute, enables value and percentage data labels formatted as percentages, sets the label font size to 12 points, and saves the chart as a PNG image.
class TimelineChartExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data: dates in column A and values in column B.
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Value");
            DateTime startDate = new DateTime(2023, 1, 1);
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(startDate.AddMonths(i));
                sheet.Cells[i + 1, 1].PutValue((i + 1) * 10); // arbitrary values
            }

            // Add a Line chart (rows 6‑20, columns 0‑10) as a substitute for Timeline.
            int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data series (values) and categories (dates).
            chart.NSeries.Add("B2:B6", true);
            chart.NSeries.CategoryData = "A2:A6";

            // Enable data labels for the series.
            // The HasDataLabel property may not be available in some versions; setting label options directly works.
            chart.NSeries[0].DataLabels.ShowValue = true;          // Show the value.
            chart.NSeries[0].DataLabels.ShowPercentage = true;    // Show percentage.

            // Format the data label to display as a percentage.
            chart.NSeries[0].DataLabels.NumberFormat = "0%";

            // Adjust the font size of the data labels.
            chart.NSeries[0].DataLabels.Font.Size = 12; // Desired font size.

            // Save the workbook as a PNG image.
            workbook.Save("TimelineChart.png", SaveFormat.Png);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
