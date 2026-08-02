// Title: Measure enumeration time of initialized cells across all worksheets in a large Aspose.Cells workbook (C#)
// Description: C# example that loads a large Excel file with Aspose.Cells, starts a Stopwatch, iterates each worksheet using Cells.GetEnumerator to count every instantiated cell, stops the timer, prints elapsed milliseconds and total cells, verifies the count with Cells.CountLarge, and saves the workbook. Ideal for benchmarking cell‑enumeration performance in .NET.
// Keywords: Aspose.Cells enumerate cells | C# cell enumeration performance | Stopwatch Aspose.Cells | large workbook benchmark | .NET Excel cell count | Cells.GetEnumerator timing | CountLarge verification | measure Excel processing speed | Aspose.Cells performance testing
// Common Searches: how to time cell enumeration in Aspose.Cells C# | benchmark initialized cell count across worksheets | measure Aspose.Cells enumeration speed for large workbooks | verify total cells with CountLarge in Aspose.Cells | C# example to log enumeration time of Excel cells
// Developer Intent: Log the duration required to enumerate every initialized cell in each worksheet of a large workbook using Aspose.Cells.
// Use Cases: Benchmarking enumeration speed to assess performance impact of large spreadsheets. | Validating that enumerated cell count matches the sum of Cells.CountLarge for data integrity. | Profiling enumeration time before and after applying optimizations such as disabling the calculation engine.
// AI Prompts: Generate C# code that uses Aspose.Cells to log enumeration time of initialized cells and compares the result with Cells.CountLarge for each worksheet. | Suggest performance‑enhancing techniques for enumerating cells in a large Aspose.Cells workbook, including API alternatives and configuration settings. | Create a unit test that asserts the cell‑enumeration time stays below a specified threshold for a given workbook size.

using System;
using System.Collections;
using System.Diagnostics;
using Aspose.Cells;

// C# example that loads a large Excel file with Aspose.Cells, starts a Stopwatch, iterates each worksheet using Cells.GetEnumerator to count every instantiated cell, stops the timer, prints elapsed milliseconds and total cells, verifies the count with Cells.CountLarge, and saves the workbook. Ideal for benchmarking cell‑enumeration performance in .NET.
class EnumerateCellsTiming
{
    static void Main()
    {
        // Load a large workbook (replace with the actual file path)
        Workbook workbook = new Workbook("LargeWorkbook.xlsx");

        // Start timing
        Stopwatch sw = Stopwatch.StartNew();

        long enumeratedCellCount = 0;

        // Enumerate initialized cells in every worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Use the Cells.GetEnumerator method (rule-provided)
            IEnumerator enumerator = sheet.Cells.GetEnumerator();
            while (enumerator.MoveNext())
            {
                // Each iteration represents an instantiated cell
                Cell cell = (Cell)enumerator.Current;
                enumeratedCellCount++;

                // (Optional) Access cell properties here if needed
                // e.g., var value = cell.Value;
            }
        }

        // Stop timing
        sw.Stop();

        Console.WriteLine($"Enumeration time: {sw.ElapsedMilliseconds} ms");
        Console.WriteLine($"Total initialized cells enumerated: {enumeratedCellCount}");

        // Verify using the CountLarge property (rule-provided)
        long sumCountLarge = 0;
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sumCountLarge += sheet.Cells.CountLarge;
        }
        Console.WriteLine($"Sum of CountLarge across worksheets: {sumCountLarge}");

        // Save the workbook (rule-provided)
        workbook.Save("LargeWorkbook_Processed.xlsx");
    }
}
