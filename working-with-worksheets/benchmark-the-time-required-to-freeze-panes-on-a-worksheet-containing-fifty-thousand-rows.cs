// Title: C# Benchmark: FreezePanes performance on a 50,000‑row worksheet with Aspose.Cells
// Description: Creates a workbook, populates 50,000 rows, warms up the FreezePanes call, measures the time to freeze 10 rows and 5 columns, and saves the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# FreezePanes benchmark | worksheet performance | 50,000 rows | .NET Excel API | measure FreezePanes latency | Excel freeze panes speed
// Common Searches: FreezePanes execution time Aspose.Cells | benchmark FreezePanes C# 50000 rows | how fast is FreezePanes on large worksheets | measure Excel freeze panes latency .NET | performance test for Aspose.Cells FreezePanes
// Developer Intent: Find out how long the FreezePanes method takes on a worksheet with 50,000 rows.
// Use Cases: Evaluate the impact of FreezePanes before adding it to a high‑volume reporting UI. | Compare freeze‑pane latency across different sheet sizes to choose optimal freeze points. | Ensure batch workbook generation remains performant when applying FreezePanes.
// AI Prompts: Write a C# function that benchmarks FreezePanes for several row counts and returns average timings. | Suggest optimization techniques to reduce FreezePanes latency on large worksheets with Aspose.Cells. | Generate a unit test that verifies FreezePanes completes within a defined time limit for a 50,000‑row sheet.

using System;
using System.Diagnostics;
using Aspose.Cells;

// Creates a workbook, populates 50,000 rows, warms up the FreezePanes call, measures the time to freeze 10 rows and 5 columns, and saves the file using Aspose.Cells for .NET.
class FreezePanesBenchmark
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate the worksheet with 50,000 rows of sample data
        for (int i = 0; i < 50000; i++)
        {
            sheet.Cells[i, 0].PutValue($"Row {i + 1}");
        }

        // Warm‑up call to avoid JIT overhead affecting the measurement
        sheet.FreezePanes(1, 1, 1, 1);
        sheet.UnFreezePanes();

        // Measure the time taken to freeze panes
        Stopwatch sw = Stopwatch.StartNew();
        // Freeze panes at row 10, column 5 with 10 frozen rows and 5 frozen columns
        sheet.FreezePanes(10, 5, 10, 5);
        sw.Stop();

        Console.WriteLine($"FreezePanes execution time: {sw.ElapsedMilliseconds} ms");

        // Save the workbook (optional, demonstrates that the workbook is still usable)
        workbook.Save("FreezePanesBenchmark.xlsx");
    }
}
