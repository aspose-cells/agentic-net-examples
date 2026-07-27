using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Charts;

class SparklineBenchmark
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        int sparklineCount = 5000;   // Number of sparklines to create
        int dataPoints = 10;         // Number of data points per sparkline

        // Populate sample data for each sparkline (rows 0..sparklineCount-1, columns 0..dataPoints-1)
        for (int row = 0; row < sparklineCount; row++)
        {
            for (int col = 0; col < dataPoints; col++)
            {
                worksheet.Cells[row, col].PutValue(col + 1); // simple incremental data
            }
        }

        // Add an empty SparklineGroup of type Line
        int groupIndex = worksheet.SparklineGroups.Add(SparklineType.Line);
        SparklineGroup sparklineGroup = worksheet.SparklineGroups[groupIndex];

        // Benchmark the creation of sparklines
        Stopwatch sw = Stopwatch.StartNew();

        for (int i = 0; i < sparklineCount; i++)
        {
            // Define the data range for the current row, e.g., "A1:J1", "A2:J2", ...
            string dataRange = $"A{i + 1}:J{i + 1}";

            // Add a sparkline at column K (index dataPoints) of the same row
            sparklineGroup.Sparklines.Add(dataRange, i, dataPoints);
        }

        sw.Stop();
        Console.WriteLine($"Created {sparklineCount} sparklines in {sw.ElapsedMilliseconds} ms.");

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("SparklineBenchmark.xlsx");
    }
}