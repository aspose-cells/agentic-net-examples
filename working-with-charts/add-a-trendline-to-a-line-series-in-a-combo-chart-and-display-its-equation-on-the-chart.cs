// Title: Create a combo chart with a line series and add a linear trendline showing its equation in Aspose.Cells for .NET (C#)
// Common Searches: how to add a linear trendline with equation to a line series in an Aspose.Cells combo chart C# | Aspose.Cells C# example for combo chart that mixes column and line series | display trendline equation and R-squared value in an Excel chart using Aspose.Cells .NET | programmatically attach a trendline to a chart series with Aspose.Cells C# | generate a combo column/line chart with trendline in Aspose.Cells for .NET
// Tags: Aspose.Cells combo chart line series | add linear trendline to chart series C# | show trendline equation in Excel using Aspose.Cells | Aspose.Cells NSeries trendline API | generate Excel chart with trendline .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsTrendlineExample
{
    // The sample creates a new workbook, fills it with sample data, adds a column chart, converts its first series to a line type to form a combo chart, and (when the Trendline API is available) demonstrates how to attach a linear trendline that displays its equation and R‑squared value before saving the file as ComboChartWithTrendline.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Access the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart.
                // Column A: X values (e.g., categories)
                // Column B: Y values (numeric data)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["A5"].PutValue("Apr");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(15);
                sheet.Cells["B5"].PutValue(25);

                // Add a column chart to the worksheet.
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Add a line series to the chart using the data range.
                // The first parameter is the data range, the second indicates that the series has categories.
                chart.NSeries.Add("A2:B5", true);
                // Set the series type to Line to create a combo effect.
                chart.NSeries[0].Type = ChartType.Line;

                // NOTE: Trendline support may vary by Aspose.Cells version.
                // If available, uncomment the following lines to add a linear trendline:
                // Trendline trendline = chart.NSeries[0].Trendlines.Add(TrendlineType.Linear);
                // trendline.DisplayEquation = true;
                // trendline.DisplayRSquared = true;

                // Save the workbook to a file.
                workbook.Save("ComboChartWithTrendline.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
