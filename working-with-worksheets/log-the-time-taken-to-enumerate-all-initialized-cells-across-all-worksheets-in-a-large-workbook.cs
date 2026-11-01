// Title: Log elapsed time for enumerating all initialized cells in every worksheet of a large Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an Excel file with Aspose.Cells, starts a Stopwatch, iterates through each worksheet's Cells collection to enumerate only populated cells, stops the timer, and prints the elapsed milliseconds. | Update an Aspose.Cells .NET example to include performance logging that measures how long it takes to traverse all populated cells across every sheet in a large workbook.
// Common Searches: how to measure performance of cell enumeration in Aspose.Cells C# | C# Aspose.Cells iterate over initialized cells and log time | benchmarking workbook traversal time with Aspose.Cells .NET | log execution time for enumerating populated cells in large Excel file using Aspose.Cells | stopwatch usage for cell enumeration across multiple worksheets Aspose.Cells
// Tags: Aspose.Cells cell traversal performance | stopwatch timing workbook scanning .NET | benchmark large Excel file cell iteration | measure worksheet traversal time | non-empty cell iteration Aspose.Cells

using System;
using System.Diagnostics;
using Aspose.Cells;

// The program loads an Excel workbook with Aspose.Cells, starts a Stopwatch, iterates through each worksheet and each initialized (populated) cell using nested foreach loops, stops the timer, and writes the elapsed milliseconds to the console.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with the actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Start timing
        Stopwatch stopwatch = Stopwatch.StartNew();

        // Enumerate all initialized cells in every worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            foreach (Cell cell in sheet.Cells)
            {
                // Enumeration only; no processing required
            }
        }

        // Stop timing
        stopwatch.Stop();

        // Log the elapsed time
        Console.WriteLine($"Time taken to enumerate all initialized cells: {stopwatch.ElapsedMilliseconds} ms");
    }
}
