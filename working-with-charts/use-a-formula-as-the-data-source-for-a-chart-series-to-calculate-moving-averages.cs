// Title: Create a line chart with a 3‑period moving average series using an array formula as the data source in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# use array formula as chart series data source | how to plot moving average line chart with Aspose.Cells .NET | set X axis values for chart series in Aspose.Cells using cell range | add line chart to worksheet and calculate 3‑point moving average in C# | Aspose.Cells formula for moving average in chart series example
// Tags: array formula chart series Aspose.Cells | 3‑point moving average line chart .NET | set chart series XValues Aspose.Cells | add line chart to worksheet C# | OFFSET AVERAGE formula Aspose.Cells chart | populate date and value columns Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// Demonstrates how to create a workbook, fill date and numeric data, add a line chart, define a 3‑point moving average using an array formula as the series source, assign the date range as X‑axis values, and save the result as MovingAverageChart.xlsx with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Access the first worksheet and name it
            Worksheet ws = wb.Worksheets[0];
            ws.Name = "Data";

            // Fill sample data: dates in column A and values in column B
            ws.Cells["A1"].PutValue("Date");
            ws.Cells["B1"].PutValue("Value");
            DateTime start = new DateTime(2023, 1, 1);
            for (int i = 0; i < 10; i++)
            {
                ws.Cells[i + 1, 0].PutValue(start.AddDays(i));   // Column A
                ws.Cells[i + 1, 1].PutValue(10 + i * 5);        // Column B
            }

            // Add a line chart to the worksheet
            int chartIndex = ws.Charts.Add(ChartType.Line, 12, 0, 26, 10);
            Chart chart = ws.Charts[chartIndex];
            chart.Title.Text = "Moving Average (3 periods)";

            // Define a formula that calculates a 3‑point moving average of column B.
            // This is an array formula that returns a range of values.
            string movingAvgFormula = "{=AVERAGE(OFFSET(B2,ROW(B2:B11)-ROW(B2),0,3,1))}";

            // Add the series using the formula as its data source.
            // The second argument 'true' indicates that the source is a formula string.
            chart.NSeries.Add(movingAvgFormula, true);

            // Set the X‑axis values (dates) for the series.
            chart.NSeries[0].XValues = "Data!A2:A11";

            // Determine output file path and ensure directory exists
            string outputPath = "MovingAverageChart.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file.
            wb.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
