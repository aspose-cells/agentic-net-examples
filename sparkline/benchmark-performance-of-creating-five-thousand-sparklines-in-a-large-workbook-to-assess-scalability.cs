// Title: Benchmarking the creation of 5,000 line sparklines in a large Excel workbook with Aspose.Cells for .NET
// AI Prompts: Write a C# program that populates a worksheet with random numbers, adds a line sparkline group, and uses Stopwatch to measure the time required to insert 5,000 sparklines. | Extend the benchmark to loop through SparklineType.Line, SparklineType.Column, and SparklineType.WinLoss, recording execution time for each type. | Create a utility that runs the sparkline insertion benchmark multiple times, aggregates the results, and outputs the average duration to evaluate scalability.
// Common Searches: aspnet benchmark adding thousands of sparklines with Aspose.Cells | measure performance of sparkline groups in large Excel files using C# | how long does it take to insert 5000 line sparklines with Aspose.Cells | scalability test for sparkline creation in .NET Excel workbook
// Tags: Aspose.Cells sparkline insertion performance | C# benchmark sparkline creation | large workbook sparkline scalability | Stopwatch timing Aspose.Cells operations | line sparkline group generation .NET

using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample generates a workbook, fills 5,000 rows with 10 columns of random numeric data, creates a line sparkline group, times the addition of one sparkline per row (total 5,000) using Stopwatch, prints the elapsed milliseconds, and saves the file as SparklineBenchmark.xlsx.
class SparklineBenchmark
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Prepare sample data: 5000 rows, 10 columns of numeric values
        const int rows = 5000;
        const int cols = 10;
        Random rnd = new Random();

        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < cols; c++)
            {
                sheet.Cells[r, c].PutValue(rnd.NextDouble() * 100);
            }
        }

        // Add a sparkline group (type Line) without initial sparklines
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Benchmark the addition of 5,000 sparklines
        Stopwatch sw = Stopwatch.StartNew();

        for (int r = 0; r < rows; r++)
        {
            // Data range for the current row (e.g., A1:J1, A2:J2, ...)
            string dataRange = $"A{r + 1}:J{r + 1}";

            // Location: column K (index 10), same row as data
            int sparklineIndex = group.Sparklines.Add(dataRange, r, cols); // cols == 10 -> column K
        }

        sw.Stop();
        Console.WriteLine($"Time to create {rows} sparklines: {sw.ElapsedMilliseconds} ms");

        // Save the workbook (output file)
        workbook.Save("SparklineBenchmark.xlsx");
    }
}
