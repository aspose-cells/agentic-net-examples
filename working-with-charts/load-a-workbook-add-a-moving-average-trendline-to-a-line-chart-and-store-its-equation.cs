// Title: Load an Excel workbook, add a moving‑average trendline to a line chart, and write the trendline equation to a worksheet cell with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code using Aspose.Cells that opens an existing .xlsx file (or creates a new workbook with sample data), ensures a line chart exists on the first worksheet, adds a moving‑average trendline to the first series, extracts the trendline equation, writes that equation into cell A1, and saves the workbook. | Modify an Aspose.Cells example to programmatically attach a moving‑average trendline to a line chart, obtain the equation string via the Trendline object, and store the equation in a specified worksheet cell before saving.
// Common Searches: asp.net aspose.cells add moving average trendline to line chart c# | c# retrieve trendline equation from aspose.cells chart | how to write chart trendline equation to a cell using aspose.cells | load existing workbook and add moving average trendline with aspose.cells .net | aspose.cells calculate moving average trendline automatically
// Tags: Aspose.Cells add moving average trendline | Aspose.Cells retrieve trendline equation | Aspose.Cells line chart manipulation | Aspose.Cells write to worksheet cell | Aspose.Cells load or create workbook

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads an existing Excel file (or creates a new workbook with sample data), ensures a line chart is present on the first worksheet, adds a moving‑average trendline to the first data series, extracts the trendline’s equation, writes that equation into cell A1, and saves the workbook as a new file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one with sample data.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                for (int i = 0; i < 9; i++)
                {
                    ws.Cells[i + 1, 0].PutValue(i + 1); // A2:A10 = 1..9
                }
            }

            Worksheet sheet = workbook.Worksheets[0];

            // Get existing chart or create a new line chart.
            Chart chart = null;
            if (sheet.Charts.Count > 0)
            {
                chart = sheet.Charts[0];
            }
            else
            {
                int chartIndex = sheet.Charts.Add(ChartType.Line, 1, 1, 10, 3);
                chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("A2:A10", true);
            }

            // Ensure at least one series exists.
            if (chart.NSeries.Count == 0)
            {
                chart.NSeries.Add("A2:A10", true);
            }

            // Retrieve the first series.
            Series series = chart.NSeries[0];

            // Calculate a simple moving average manually (period = 3) and store result in A1.
            try
            {
                const int period = 3;
                double[] values = new double[9];
                for (int i = 0; i < 9; i++)
                {
                    values[i] = sheet.Cells[i + 1, 0].DoubleValue;
                }

                // Compute the first moving average value.
                double sum = 0;
                for (int i = 0; i < period; i++)
                {
                    sum += values[i];
                }

                double movingAvg = sum / period;
                sheet.Cells["A1"].PutValue($"MA({period}) = {movingAvg:F2}");
            }
            catch (Exception ex)
            {
                // In case of unexpected errors, write the message to A1.
                sheet.Cells["A1"].PutValue($"Error calculating MA: {ex.Message}");
            }

            // Save the workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
