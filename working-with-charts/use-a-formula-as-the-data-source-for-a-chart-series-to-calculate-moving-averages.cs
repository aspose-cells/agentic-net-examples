// Title: Generate a line chart with a 3‑point moving average series using formula‑driven cells in Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to fill column A with sample values, insert AVERAGE formulas in column B for a 3‑point moving average, and create a line chart that plots both the original data and the calculated moving‑average series. | Demonstrate how to bind a range containing Excel formulas as the data source for a chart series in Aspose.Cells, set series names, and save the workbook.
// Common Searches: how to add a moving average series to a chart with Aspose.Cells C# | Aspose.Cells line chart using formula cells as data source | C# calculate 3 point moving average with AVERAGE formula in Aspose.Cells | bind calculated range to chart series Aspose.Cells .NET example | create chart with original data and moving average in Aspose.Cells workbook
// Tags: Aspose.Cells line chart with formula data source | C# moving average calculation using AVERAGE formula | bind calculated cells to chart series Aspose.Cells | create Excel chart with original and moving average data | Aspose.Cells chart series from computed range

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// The example creates a workbook, writes sample numbers to column A, adds 3‑point moving‑average formulas to column B, then builds a line chart that displays the raw data (A1:A10) and the computed moving‑average series (B3:B10). The chart is titled, series are named, and the workbook is saved as MovingAverageChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Fill sample data in column A (A1:A10)
            double[] data = { 10, 12, 15, 14, 13, 16, 18, 20, 19, 22 };
            for (int i = 0; i < data.Length; i++)
            {
                cells[i, 0].PutValue(data[i]); // Row i, Column 0 (A)
            }

            // Insert moving‑average formulas in column B (B3:B10) – 3‑point moving average
            for (int row = 2; row < data.Length; row++) // start from third data point (row index 2)
            {
                // Excel formula: =AVERAGE(A{row}:A{row+2})
                string formula = $"AVERAGE(A{row + 1}:A{row + 3})";
                cells[row, 1].Formula = formula; // Column B
            }

            // Add a line chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Original Data vs. 3‑Point Moving Average";

            // Series 1 – original data (A1:A10)
            chart.NSeries.Add("A1:A10", true);
            chart.NSeries[0].Name = "Original Data";

            // Series 2 – moving average (B3:B10)
            chart.NSeries.Add("B3:B10", true);
            chart.NSeries[1].Name = "3‑Point Moving Avg";

            // Save the workbook
            workbook.Save("MovingAverageChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
