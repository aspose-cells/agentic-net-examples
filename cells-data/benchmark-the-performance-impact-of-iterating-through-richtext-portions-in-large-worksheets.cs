// Title: Benchmarking rich‑text character iteration performance in large Aspose.Cells worksheets using C#
// Common Searches: C# benchmark iterating cell.Characters across thousands of rows in Aspose.Cells | measure performance of rich‑text fragment access in a large Excel workbook with Aspose.Cells | how to time character‑level rich‑text processing in Aspose.Cells worksheets | performance test for accessing colored text fragments in Aspose.Cells using Stopwatch | evaluate iteration speed of cell.Characters method on a 20,000‑row worksheet
// Tags: rich text character iteration Aspose.Cells | cell.Characters performance measurement C# | large worksheet processing Aspose.Cells | Aspose.Cells rich text access timing | MaxDisplayRange iteration optimization Aspose.Cells | colored text fragment iteration Aspose.Cells

using System;
using System.Diagnostics;
using System.Drawing;
using Aspose.Cells;

// The example creates a workbook with 20,000 rows and 5 columns, fills each cell with the string "AlphaBetaGamma" and applies three colored fragments via the Characters API. After a warm‑up call, it iterates through every character of every populated cell, accesses each character's Font.Color to simulate work, measures the elapsed time with Stopwatch, and outputs the total cells processed and the duration, providing a benchmark of rich‑text character iteration in large worksheets.
class RichTextIterationBenchmark
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the size of the test data
            const int totalRows = 20000;   // large number of rows
            const int totalCols = 5;       // number of columns per row

            // Populate the worksheet with rich‑text cells
            for (int row = 0; row < totalRows; row++)
            {
                for (int col = 0; col < totalCols; col++)
                {
                    // Build a rich string with three colored fragments
                    string richText = "AlphaBetaGamma";
                    Cell cell = sheet.Cells[row, col];
                    cell.PutValue(richText);

                    // Apply colors to the fragments using the Characters method
                    cell.Characters(0, 5).Font.Color = Color.Red;      // "Alpha"
                    cell.Characters(5, 4).Font.Color = Color.Green;    // "Beta"
                    cell.Characters(9, 5).Font.Color = Color.Blue;     // "Gamma"
                }
            }

            // Warm‑up: ensure any lazy initialization is done before measurement
            workbook.CalculateFormula();

            // Benchmark: iterate through all rich‑text portions (character by character)
            Stopwatch sw = Stopwatch.StartNew();

            // Use the used range to limit iteration to populated cells
            Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

            for (int row = usedRange.FirstRow; row <= usedRange.FirstRow + usedRange.RowCount - 1; row++)
            {
                for (int col = usedRange.FirstColumn; col <= usedRange.FirstColumn + usedRange.ColumnCount - 1; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    string text = cell.StringValue;

                    // Iterate over each character to simulate fragment processing
                    for (int i = 0; i < text.Length; i++)
                    {
                        // Access character properties (e.g., color) to simulate work
                        Color color = cell.Characters(i, 1).Font.Color;
                        // No further processing needed for the benchmark
                    }
                }
            }

            sw.Stop();

            // Output the elapsed time
            Console.WriteLine($"Iterated through rich‑text portions in {totalRows * totalCols} cells.");
            Console.WriteLine($"Elapsed time: {sw.Elapsed.TotalSeconds:F3} seconds.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
