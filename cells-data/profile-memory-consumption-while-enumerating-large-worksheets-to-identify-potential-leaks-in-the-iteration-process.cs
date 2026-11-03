// Title: How to profile memory usage while iterating a large Excel worksheet with Aspose.Cells for .NET
// Common Searches: C# measure memory consumption during Aspose.Cells worksheet iteration | detect memory leaks while looping through cells of a large .xlsx using Aspose.Cells | profile per‑row memory usage in a big Excel file with Aspose.Cells .NET | log GC.GetTotalMemory after processing each row in Aspose.Cells | benchmark memory impact of accessing cell values in Aspose.Cells
// Tags: Aspose.Cells worksheet iteration memory profiling | C# GC memory measurement per row Excel | large .xlsx processing memory analysis Aspose.Cells | detect memory leaks Aspose.Cells cell enumeration | benchmark cell value access memory impact .NET

using System;
using System.Diagnostics;
using Aspose.Cells;

// The example loads a large Excel workbook with Aspose.Cells, forces garbage collection to set a baseline, then iterates through every used cell while accessing each cell's Value. After each row it records the current managed memory using GC.GetTotalMemory, allowing developers to observe per‑row memory growth and identify potential leaks during worksheet enumeration.
class Program
{
    static void Main()
    {
        // Load the workbook containing a large worksheet
        Workbook workbook = new Workbook("LargeData.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Ensure a clean baseline by forcing garbage collection
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();

        long initialMemory = GC.GetTotalMemory(true);
        Console.WriteLine($"Initial memory: {initialMemory / 1024 / 1024} MB");

        // Determine the used range of the worksheet
        int maxRow = cells.MaxDataRow;
        int maxColumn = cells.MaxDataColumn;

        // Iterate through each row and column, accessing cell values
        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxColumn; col++)
            {
                // Access the cell to simulate processing work
                Cell cell = cells[row, col];
                var value = cell.Value;
            }

            // Capture memory usage after processing each row
            long currentMemory = GC.GetTotalMemory(false);
            Console.WriteLine($"After row {row}: {currentMemory / 1024 / 1024} MB");
        }

        // Final memory measurement after iteration completes
        long finalMemory = GC.GetTotalMemory(true);
        Console.WriteLine($"Final memory: {finalMemory / 1024 / 1024} MB");
    }
}
