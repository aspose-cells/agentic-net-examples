// Title: Measure and log the time required to enumerate all initialized cells in every worksheet of a large Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens a workbook with Aspose.Cells, iterates through each worksheet's Cells collection, accesses each cell value, and records the elapsed time with System.Diagnostics.Stopwatch. | Show how to profile the enumeration of populated cells across all sheets in a large Excel file by printing the total milliseconds taken.
// Common Searches: how to time cell enumeration across all sheets with Aspose.Cells in C# | performance test for iterating initialized cells in a large Excel workbook using Aspose.Cells | C# measure elapsed milliseconds when looping through populated cells in each worksheet
// Tags: enumerate initialized cells Aspose.Cells .NET | stopwatch performance measurement Aspose.Cells | large workbook cell iteration timing | benchmark cell enumeration Aspose.Cells | worksheet cells enumeration performance

using System;
using System.Diagnostics;
using Aspose.Cells;

// Loads a workbook, loops through every worksheet's Cells collection (which contains only initialized cells), accesses each cell value to force enumeration, and prints the elapsed time in milliseconds using Stopwatch.
class Program
{
    static void Main()
    {
        // Load the large workbook
        Workbook workbook = new Workbook("LargeWorkbook.xlsx");

        // Start timing
        Stopwatch stopwatch = Stopwatch.StartNew();

        // Enumerate all initialized cells in every worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // The Cells collection enumerates only cells that have data, formulas, or styles applied
            foreach (Cell cell in sheet.Cells)
            {
                // Accessing the cell forces enumeration; no further processing required
                var _ = cell.Value;
            }
        }

        // Stop timing
        stopwatch.Stop();

        // Log the elapsed time
        Console.WriteLine($"Time taken to enumerate all initialized cells: {stopwatch.ElapsedMilliseconds} ms");
    }
}
