// Title: Generate a line chart with a 3‑period moving average series using an AVERAGE formula as the data source in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set chart series values with AVERAGE formula | How to calculate moving average for a chart series in Aspose.Cells .NET | Assign date range as X‑axis categories for a line chart using Aspose.Cells | Create line chart with dynamic data source formula in C# Aspose.Cells | Programmatically add moving average series to an Excel chart with Aspose.Cells
// Tags: chart series formula source Aspose.Cells | 3‑period moving average chart .NET | assign X axis range Aspose.Cells | line chart generation C# Aspose.Cells | AVERAGE formula as series values Excel

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example creates a new workbook, fills column A with sequential dates and column B with incremental numbers, adds a line chart, defines a series whose values are calculated by the AVERAGE formula over a three‑row range to produce a moving average, sets the dates range as the X‑axis categories, and saves the file as MovingAverageChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Get the first worksheet and name it
            Worksheet ws = wb.Worksheets[0];
            ws.Name = "Data";

            // Fill sample data: dates in column A and numeric values in column B
            for (int i = 0; i < 10; i++)
            {
                ws.Cells[i, 0].PutValue(DateTime.Today.AddDays(i)); // A column (dates)
                ws.Cells[i, 1].PutValue(10 + i * 2);               // B column (values)
            }

            // Add a line chart to the worksheet (positioned below the data)
            int chartIdx = ws.Charts.Add(ChartType.Line, 12, 0, 26, 10);
            Chart chart = ws.Charts[chartIdx];
            chart.Title.Text = "3‑Period Moving Average";

            // Add a series whose values are defined by a formula.
            // The formula calculates a 3‑period moving average of column B.
            int seriesIdx = chart.NSeries.Add("=AVERAGE(B1:B3)", true);
            Series series = chart.NSeries[seriesIdx];

            // Use the dates column as the X‑axis (category) values
            series.XValues = "A1:A10";

            // Ensure the output directory exists
            string outputPath = "MovingAverageChart.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
