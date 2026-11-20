// Title: Benchmarking iteration over simulated rich‑text portions in large Excel worksheets using Aspose.Cells for .NET (C#)
// Common Searches: how to benchmark cell iteration performance with Aspose.Cells in C# | measure time to split simulated rich‑text portions in Excel cells using Aspose.Cells | C# performance test for iterating 20,000 rows with multiple text portions per cell | Aspose.Cells benchmark for processing rich‑text‑like strings in large worksheets | stopwatch timing of cell string split operations in Aspose.Cells .NET
// Tags: Aspose.Cells rich text iteration performance | C# benchmark large worksheet cell processing | simulate multiple text portions per Excel cell | cell string split latency measurement Aspose.Cells | performance testing Aspose.Cells cell iteration

using System;
using System.Diagnostics;
using Aspose.Cells;

// The example creates a workbook with 20,000 rows where each cell in column A contains a space‑separated string of five parts to mimic rich‑text portions. After a warm‑up pass, it uses a Stopwatch to benchmark the time required to iterate through all cells, split each cell's string into its simulated portions, and compute the length of each part, demonstrating the performance impact of rich‑text‑like iteration in large worksheets.
class RichTextIterationBenchmark
{
    static void Main()
    {
        try
        {
            // Initialize a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Parameters for the test
            int totalRows = 20000;          // Number of rows to create
            int richTextPortions = 5;       // Number of simulated rich‑text portions per cell

            // Populate the worksheet with simulated rich‑text cells
            for (int row = 0; row < totalRows; row++)
            {
                // Build a plain string that represents multiple portions
                string cellText = string.Empty;
                for (int i = 0; i < richTextPortions; i++)
                {
                    cellText += $"Part{i} ";
                }

                // Assign the string to the cell
                cells[row, 0].PutValue(cellText.TrimEnd());
            }

            // Warm‑up: iterate once to avoid JIT overhead affecting the measurement
            IterateRichText(cells, totalRows, richTextPortions);

            // Benchmark the iteration over simulated rich‑text portions
            Stopwatch sw = Stopwatch.StartNew();
            IterateRichText(cells, totalRows, richTextPortions);
            sw.Stop();

            Console.WriteLine($"Iterated {totalRows} cells with {richTextPortions} simulated portions each in {sw.ElapsedMilliseconds} ms.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper method that iterates through simulated rich‑text portions of cells in column A
    static void IterateRichText(Cells cells, int rows, int portionsPerCell)
    {
        for (int row = 0; row < rows; row++)
        {
            Cell cell = cells[row, 0];
            // Retrieve the cell value as a string
            string text = cell.StringValue ?? string.Empty;

            // Simulate portion iteration by splitting on spaces
            string[] parts = text.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string part in parts)
            {
                // Dummy operation: compute length (no side effects)
                int len = part.Length;
            }
        }
    }
}
