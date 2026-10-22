// Title: Record and log the elapsed time of each Worksheet.FreezePanes call with Stopwatch in Aspose.Cells for .NET
// AI Prompts: Write C# code that wraps Worksheet.FreezePanes in a Stopwatch, prints the elapsed milliseconds, and repeats for multiple freeze configurations. | Create a helper method in Aspose.Cells that accepts freeze parameters, measures the call duration with System.Diagnostics.Stopwatch, and returns the timing value for logging.
// Common Searches: how to benchmark Worksheet.FreezePanes performance using Stopwatch in C# | Aspose.Cells log time taken by each FreezePanes invocation | measure execution latency of freeze panes in .NET workbook | C# example for timing multiple FreezePanes calls with Aspose.Cells | record FreezePanes call duration and output to console Aspose
// Tags: Aspose.Cells FreezePanes execution timing | C# Stopwatch worksheet freeze panes | benchmark FreezePanes performance Aspose | log FreezePanes latency .NET | measure worksheet freeze panes duration

using Aspose.Cells;
using System;
using System.Diagnostics;

// // Demonstrates using Stopwatch to record and output the elapsed milliseconds for three separate Worksheet.FreezePanes calls, then saves the workbook as FreezePanesDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // -----------------------------------------------------------------
        // First FreezePanes call: freeze the first row
        // -----------------------------------------------------------------
        Stopwatch sw = Stopwatch.StartNew();               // Start timing
        sheet.FreezePanes(1, 0, 1, 0);                     // Freeze first row
        sw.Stop();                                         // Stop timing
        Console.WriteLine($"FreezePanes call 1 elapsed: {sw.ElapsedMilliseconds} ms");

        // -----------------------------------------------------------------
        // Second FreezePanes call: freeze the first column
        // -----------------------------------------------------------------
        sw.Restart();                                      // Restart timing
        sheet.FreezePanes(0, 1, 0, 1);                     // Freeze first column
        sw.Stop();                                         // Stop timing
        Console.WriteLine($"FreezePanes call 2 elapsed: {sw.ElapsedMilliseconds} ms");

        // -----------------------------------------------------------------
        // Additional FreezePanes call example: freeze first row and column
        // -----------------------------------------------------------------
        sw.Restart();                                      // Restart timing
        sheet.FreezePanes(1, 1, 1, 1);                     // Freeze first row and column
        sw.Stop();                                         // Stop timing
        Console.WriteLine($"FreezePanes call 3 elapsed: {sw.ElapsedMilliseconds} ms");

        // Save the workbook to a file
        workbook.Save("FreezePanesDemo.xlsx");
    }
}
