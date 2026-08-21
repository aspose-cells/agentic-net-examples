// Title: Benchmark: Time to Enumerate All Initialized Cells in a Large Aspose.Cells Workbook (C#)
// Description: This example creates (or loads) a workbook, populates thousands of cells, starts a Stopwatch, iterates through every worksheet, uses the Cells enumerator to walk through each instantiated cell, counts them, stops the timer, and writes the elapsed seconds and total cell count to the console. The workbook can then be saved.
// Keywords: Aspose.Cells | C# | enumerate initialized cells | cell enumeration performance | Stopwatch benchmark | large workbook | .NET Excel processing | count populated cells | performance testing | Excel cell traversal
// Common Searches: Aspose.Cells enumerate initialized cells performance | measure time to iterate over all cells in a workbook C# | benchmark cell enumeration across worksheets Aspose | how to count populated cells in a large Excel file using Aspose.Cells | stopwatch cell traversal Aspose.Cells .NET
// Developer Intent: Find out how long it takes to traverse and count every initialized cell in all worksheets of a large Aspose.Cells workbook.
// Use Cases: Benchmarking cell‑traversal speed to gauge performance impact of large data imports. | Verifying that all populated cells are reachable before applying further processing or saving. | Estimating execution time for custom validation or transformation logic that iterates over every initialized cell.
// AI Prompts: Generate C# code that uses Aspose.Cells to enumerate only non‑empty cells in each worksheet, count them, and output the elapsed time. | Suggest a more efficient method to obtain the total number of initialized cells in a large workbook without iterating through each cell individually. | Provide an example that logs enumeration performance to a file instead of the console, using Aspose.Cells and Stopwatch.

using System;
using System.Collections;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example creates (or loads) a workbook, populates thousands of cells, starts a Stopwatch, iterates through every worksheet, uses the Cells enumerator to walk through each instantiated cell, counts them, stops the timer, and writes the elapsed seconds and total cell count to the console. The workbook can then be saved.
    public class EnumerateInitializedCellsDemo
    {
        public static void Main()
        {
            // Create a new workbook (or load an existing large workbook)
            // For demonstration, we create a workbook and populate many cells.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate a large number of cells to simulate a large workbook
            // (In real scenario, replace this with workbook = new Workbook("LargeFile.xlsx");)
            for (int row = 0; row < 5000; row++)
            {
                for (int col = 0; col < 50; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row}C{col}");
                }
            }

            // Start timing the enumeration of all initialized cells across all worksheets
            Stopwatch sw = Stopwatch.StartNew();

            long totalInitializedCells = 0;

            // Iterate through each worksheet in the workbook
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Get the Cells collection for the current worksheet
                Cells cells = ws.Cells;

                // Obtain the enumerator for the cells collection
                IEnumerator enumerator = cells.GetEnumerator();

                // Enumerate all instantiated cells
                while (enumerator.MoveNext())
                {
                    // The current object is a Cell; we can cast if needed
                    // Cell cell = (Cell)enumerator.Current;
                    totalInitializedCells++;
                }
            }

            sw.Stop();

            // Log the results
            Console.WriteLine($"Time taken to enumerate initialized cells: {sw.Elapsed.TotalSeconds:F3} seconds");
            Console.WriteLine($"Total initialized cells enumerated: {totalInitializedCells}");

            // Optionally save the workbook (demonstrates usage of save lifecycle)
            workbook.Save("EnumeratedWorkbook.xlsx");
        }
    }
}
