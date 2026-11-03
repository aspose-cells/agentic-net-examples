// Title: How to profile CPU usage while enumerating all cells in a large worksheet using Aspose.Cells for .NET
// Common Searches: c# profile CPU usage when looping through cells with Aspose.Cells | measure execution time of worksheet cell enumeration Aspose.Cells .NET | detect performance bottlenecks in large Excel workbook processing using Aspose.Cells | use Process.TotalProcessorTime to monitor CPU while reading Excel cells in C# | benchmark cell read loop performance for a large worksheet with Aspose.Cells
// Tags: cell enumeration CPU profiling Aspose.Cells | worksheet loop processor time measurement | Aspose.Cells large workbook iteration benchmarking | wall‑clock vs CPU time measurement C# | performance analysis of Excel cell traversal

using System;
using System.Diagnostics;
using Aspose.Cells;

// The example loads a workbook, determines the used row and column range, and then measures both CPU time (via Process.TotalProcessorTime) and wall‑clock time (via Stopwatch) while nested loops read every cell's value. After the enumeration it prints the total rows, columns, elapsed wall‑clock seconds, CPU seconds, and the CPU usage ratio, helping identify bottlenecks in large worksheet iteration.
class WorksheetEnumerationProfiler
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("LargeData.xlsx");

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Determine the used range
        int maxRow = cells.MaxDataRow;      // zero‑based index of the last row with data
        int maxCol = cells.MaxDataColumn;   // zero‑based index of the last column with data

        // Prepare performance measurement tools
        Process currentProcess = Process.GetCurrentProcess();
        Stopwatch wallClock = new Stopwatch();

        // Capture initial CPU time
        TimeSpan cpuStart = currentProcess.TotalProcessorTime;

        // Start wall‑clock timing
        wallClock.Start();

        // Enumerate the worksheet cells
        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                // Access the cell (this is the operation we want to profile)
                Cell cell = cells[row, col];
                // Example read – you can replace with any processing logic
                var value = cell.Value;
            }
        }

        // Stop wall‑clock timing
        wallClock.Stop();

        // Capture final CPU time
        TimeSpan cpuEnd = currentProcess.TotalProcessorTime;

        // Calculate elapsed times
        TimeSpan cpuElapsed = cpuEnd - cpuStart;
        TimeSpan wallElapsed = wallClock.Elapsed;

        // Output profiling results
        Console.WriteLine("Worksheet enumeration completed.");
        Console.WriteLine($"Rows processed: {maxRow + 1}");
        Console.WriteLine($"Columns processed: {maxCol + 1}");
        Console.WriteLine($"Wall‑clock time: {wallElapsed.TotalSeconds:F3} seconds");
        Console.WriteLine($"CPU time used: {cpuElapsed.TotalSeconds:F3} seconds");
        Console.WriteLine($"CPU usage ratio: {(cpuElapsed.TotalSeconds / wallElapsed.TotalSeconds):P2}");

        // Optionally, save the workbook if any modifications were made
        // workbook.Save("LargeData_Processed.xlsx");
    }
}
