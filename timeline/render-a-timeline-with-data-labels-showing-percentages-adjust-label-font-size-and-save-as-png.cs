// Title: Create a timeline line chart with percentage data labels, adjust label font size, and export as PNG using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# line chart with date axis and percentage data labels | Set custom font size for chart data labels using Aspose.Cells .NET | Export a chart as PNG image with Aspose.Cells in C# | Render a project timeline chart with values and percentages in Aspose.Cells | How to enable both value and percentage in Aspose.Cells chart labels
// Tags: timeline line chart generation Aspose.Cells | adjust label typography Aspose.Cells chart | chart image export PNG Aspose.Cells | axis formatting for dates Aspose.Cells | percentage display in chart labels Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, fills it with dates and numeric values, adds a line chart as a project timeline, enables data labels to show both actual values and their percentages, sets the label font size to 12 points, and saves the rendered chart as a PNG image.
class TimelineChartExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data: dates in column A and values in column B.
            // Row 0 is the header.
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Value");

            // Sample dates and values.
            DateTime[] dates = {
                new DateTime(2023, 1, 1),
                new DateTime(2023, 2, 1),
                new DateTime(2023, 3, 1),
                new DateTime(2023, 4, 1),
                new DateTime(2023, 5, 1)
            };
            double[] values = { 10, 20, 30, 25, 15 };

            // Fill the worksheet with the sample data.
            for (int i = 0; i < dates.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(dates[i]);   // Column A
                sheet.Cells[i + 1, 1].PutValue(values[i]); // Column B
            }

            // Add a Line chart (ChartType.Line) to the worksheet.
            // Position the chart from row 7, column 0 to row 27, column 7.
            int chartIndex = sheet.Charts.Add(ChartType.Line, 7, 0, 27, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Set the chart title.
            chart.Title.Text = "Project Timeline";

            // Add a series to the chart using the date column as X values and the value column as Y values.
            // X values (dates) are in A2:A6, Y values (values) are in B2:B6.
            chart.NSeries.Add("B2:B6", true);
            chart.NSeries[0].XValues = "A2:A6";

            // Enable data labels.
            chart.NSeries[0].DataLabels.ShowValue = true;          // Show the actual value.
            chart.NSeries[0].DataLabels.ShowPercentage = true;    // Show percentage of total.

            // Adjust the font size of the data labels.
            chart.NSeries[0].DataLabels.Font.Size = 12; // Set desired font size (e.g., 12 points).

            // Optional: format the date axis to display month/year.
            // The NumberFormat property may not be available in all versions; skip if unsupported.
            // chart.CategoryAxis.NumberFormat = "mmm yyyy";

            // Save the workbook as a PNG image. The chart will be rendered as an image in the output file.
            workbook.Save("TimelineChart.png", SaveFormat.Png);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
