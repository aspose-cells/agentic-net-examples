// Title: Benchmark Rich‑Text Segment Iteration in Large Worksheets with Aspose.Cells for .NET
// Description: Creates a 1,000 × 10 worksheet, fills each cell with a three‑part rich‑text string (bold, italic, red), then uses a Stopwatch to measure the time required to read the Font properties of each segment via the Characters API. The elapsed milliseconds are printed and the workbook is saved.
// Keywords: Aspose.Cells | .NET | rich text iteration | Characters API performance | cell formatting benchmark | large worksheet speed test | Excel rich‑text processing | performance measurement | bulk cell iteration | Aspose.Cells performance
// Common Searches: Aspose.Cells benchmark reading rich‑text formatting | how fast is Characters API in large Excel sheets | measure performance of rich‑text iteration Aspose.Cells .NET | speed test for iterating formatted text in thousands of cells | performance impact of cell.Characters in Aspose.Cells
// Developer Intent: Quantify the execution time needed to access rich‑text formatting for each segment across a high‑volume cell range.
// Use Cases: Assess feasibility of bulk rich‑text scanning before applying transformations. | Compare processing cost of styled text versus plain text when extracting data. | Evaluate the benefit of parallel loops or calculation disabling on iteration speed.
// AI Prompts: Provide a Parallel.For implementation that speeds up the rich‑text iteration loop. | Show how to log per‑cell timing and generate a summary of fastest and slowest cells. | Suggest memory‑efficient techniques for accessing Characters objects in a large workbook.

using System;
using System.Diagnostics;
using Aspose.Cells;
using System.Drawing;

// Creates a 1,000 × 10 worksheet, fills each cell with a three‑part rich‑text string (bold, italic, red), then uses a Stopwatch to measure the time required to read the Font properties of each segment via the Characters API. The elapsed milliseconds are printed and the workbook is saved.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Define the size of the test sheet
        int totalRows = 1000;   // number of rows
        int totalCols = 10;     // number of columns

        // Populate each cell with a rich‑text string that has three formatted parts
        // Part1 : bold, Part2 : italic, Part3 : red color
        for (int row = 0; row < totalRows; row++)
        {
            for (int col = 0; col < totalCols; col++)
            {
                // Set the cell value
                cells[row, col].PutValue("Part1 Part2 Part3");

                // Apply formatting to each part
                // "Part1"
                cells[row, col].Characters(0, 5).Font.IsBold = true;

                // "Part2"
                cells[row, col].Characters(6, 5).Font.IsItalic = true;

                // "Part3"
                cells[row, col].Characters(12, 5).Font.Color = Color.Red;
            }
        }

        // Start timing the iteration over rich‑text portions
        Stopwatch sw = Stopwatch.StartNew();

        // Iterate through every cell and read the formatting of each rich‑text segment
        for (int row = 0; row < totalRows; row++)
        {
            for (int col = 0; col < totalCols; col++)
            {
                // Retrieve the cell
                Cell cell = cells[row, col];

                // Access each formatted segment (the indices match the ones used during creation)
                var part1 = cell.Characters(0, 5);   // bold segment
                var part2 = cell.Characters(6, 5);   // italic segment
                var part3 = cell.Characters(12, 5);  // red segment

                // Dummy reads to simulate work and prevent compiler optimizations
                bool isBold = part1.Font.IsBold;
                bool isItalic = part2.Font.IsItalic;
                Color color = part3.Font.Color;
            }
        }

        sw.Stop();
        Console.WriteLine($"Iterating rich‑text portions in {totalRows * totalCols} cells took {sw.ElapsedMilliseconds} ms.");

        // Save the workbook (optional, demonstrates normal save flow)
        workbook.Save("RichTextBenchmark.xlsx");
    }
}
