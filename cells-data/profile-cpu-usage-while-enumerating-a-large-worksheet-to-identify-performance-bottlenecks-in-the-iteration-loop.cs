// Title: Measure CPU and wall‑clock time while iterating over every cell in a large Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Add Process.TotalProcessorTime and Stopwatch measurements around a nested cell loop to profile CPU usage in Aspose.Cells. | Write a C# snippet that logs CPU time for each iteration while reading cell values from a large worksheet. | Modify the enumeration code to output both wall‑clock and CPU elapsed times for performance analysis.
// Common Searches: how to benchmark cell iteration performance with Aspose.Cells in C# | C# measure CPU time while reading all cells from a large Excel file using Aspose.Cells | profile worksheet enumeration speed Aspose.Cells .NET large dataset | detect performance bottleneck in Aspose.Cells cell loop
// Tags: cpu profiling Aspose.Cells cell enumeration | wall‑clock timing Aspose.Cells worksheet loop | performance measurement Aspose.Cells .NET | used range iteration Aspose.Cells | benchmark reading Excel cells C#

using System;
using System.Diagnostics;
using Aspose.Cells;

// The example loads a workbook, determines the used range, then iterates over every cell while measuring wall‑clock time with Stopwatch and CPU time via Process.TotalProcessorTime, finally printing both elapsed durations to help identify performance bottlenecks.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("LargeData.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Determine the used range
        int maxRow = cells.MaxDataRow;
        int maxColumn = cells.MaxDataColumn;

        // Get the current process to measure CPU usage
        Process currentProcess = Process.GetCurrentProcess();

        // Record CPU time before enumeration
        TimeSpan cpuStart = currentProcess.TotalProcessorTime;
        Stopwatch wallClock = Stopwatch.StartNew();

        // Enumerate all cells in the used range
        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxColumn; col++)
            {
                // Example operation: read the cell value
                object value = cells[row, col].Value;

                // Minimal processing to avoid affecting measurement
                if (value != null && value is string s && s.Length > 0)
                {
                    // No-op placeholder
                }
            }
        }

        // Stop timers
        wallClock.Stop();
        TimeSpan cpuEnd = currentProcess.TotalProcessorTime;

        // Calculate elapsed times
        TimeSpan cpuElapsed = cpuEnd - cpuStart;

        // Output profiling results
        Console.WriteLine($"Wall‑clock time: {wallClock.Elapsed.TotalSeconds:F2} seconds");
        Console.WriteLine($"CPU time used: {cpuElapsed.TotalSeconds:F2} seconds");
    }
}
