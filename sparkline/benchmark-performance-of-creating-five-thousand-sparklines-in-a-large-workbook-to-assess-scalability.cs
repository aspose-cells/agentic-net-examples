// Title: Aspose.Cells .NET Benchmark: Create 5,000 Line Sparklines in a Large Workbook
// Description: C# sample that builds a workbook, fills 5,000 rows with ten random numbers each, adds a line sparkline group, measures the time to insert a sparkline per row, outputs the elapsed seconds, and optionally saves the file—demonstrating scalability of bulk sparkline creation with Aspose.Cells.
// Keywords: Aspose.Cells | .NET | sparkline benchmark | 5000 sparklines | performance testing | large workbook | line sparkline | Excel scalability | measure creation time | bulk sparkline insertion
// Common Searches: Aspose.Cells sparkline performance test | benchmark adding thousands of sparklines in C# | how fast can Aspose.Cells create 5000 sparklines | measure sparkline creation time .NET | scalability of sparkline groups in large Excel files | performance of line sparklines with Aspose.Cells | optimize bulk sparkline insertion Aspose.Cells
// Developer Intent: Evaluate the time required to add 5,000 line sparklines and gauge Aspose.Cells scalability for bulk operations.
// Use Cases: Run the benchmark to verify that bulk sparkline insertion meets latency targets before generating production reports. | Swap SparklineType (Column, WinLoss) to compare creation speed across different visualizations. | Integrate the timing output into CI pipelines to detect performance regressions in workbook generation.
// AI Prompts: Generate a memory‑efficient version of the benchmark that uses streaming or reduced object allocation while preserving accurate timing. | Provide thread‑safe code to parallelize the addition of 5,000 sparklines with Aspose.Cells without risking workbook corruption. | Explain how to interpret the elapsed seconds and suggest tuning steps (e.g., disabling calculation, reusing ranges) to improve performance for massive Excel files.

using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Charts;

// C# sample that builds a workbook, fills 5,000 rows with ten random numbers each, adds a line sparkline group, measures the time to insert a sparkline per row, outputs the elapsed seconds, and optionally saves the file—demonstrating scalability of bulk sparkline creation with Aspose.Cells.
class SparklineBenchmark
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Parameters for the benchmark
        const int sparklineCount = 5000;   // number of sparklines to create
        const int dataColumns = 10;        // number of columns in each data range

        // Populate the worksheet with sample data for each sparkline
        Random rnd = new Random();
        for (int row = 0; row < sparklineCount; row++)
        {
            for (int col = 0; col < dataColumns; col++)
            {
                sheet.Cells[row, col].PutValue(rnd.NextDouble() * 100);
            }
        }

        // Add a sparkline group (type Line) without initial sparklines
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Benchmark the creation of 5,000 sparklines
        Stopwatch sw = Stopwatch.StartNew();

        for (int row = 0; row < sparklineCount; row++)
        {
            // Build the data range string for the current row, e.g., "A1:J1"
            string startCell = CellsHelper.CellIndexToName(row, 0);
            string endCell = CellsHelper.CellIndexToName(row, dataColumns - 1);
            string dataRange = $"{startCell}:{endCell}";

            // Location column for the sparkline (placed after the data columns)
            int locationColumn = dataColumns; // e.g., column K (index 10)

            // Add the sparkline to the group
            group.Sparklines.Add(dataRange, row, locationColumn);
        }

        sw.Stop();
        Console.WriteLine($"Created {sparklineCount} sparklines in {sw.Elapsed.TotalSeconds:F2} seconds.");

        // Save the workbook (optional, can be omitted if only benchmarking)
        workbook.Save("SparklineBenchmark.xlsx");
    }
}
