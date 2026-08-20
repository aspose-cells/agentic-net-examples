// Title: Benchmark FreezePanes Calls in Aspose.Cells for .NET Using Stopwatch
// Description: C# sample that creates a workbook, times two FreezePanes overloads with System.Diagnostics.Stopwatch, outputs the elapsed milliseconds to the console, and saves the file as FreezePanesTiming.xlsx.
// Keywords: Aspose.Cells | FreezePanes performance | Stopwatch timing | C# worksheet benchmark | execution time measurement | log FreezePanes duration | .NET performance testing | worksheet freeze panes latency
// Common Searches: how to time FreezePanes in Aspose.Cells | measure FreezePanes execution C# | benchmark worksheet FreezePanes method | record FreezePanes latency .NET | Aspose.Cells performance testing example
// Developer Intent: Determine the runtime of each FreezePanes overload and capture the results for performance analysis.
// Use Cases: Compare the speed of different FreezePanes signatures on large spreadsheets. | Detect performance regressions when updating Aspose.Cells versions. | Integrate FreezePanes timing into automated build or CI pipelines.
// AI Prompts: Generate a reusable C# method that wraps any worksheet.FreezePanes call and returns the elapsed milliseconds. | Show how to export FreezePanes timing results to a CSV file for later analysis. | Explain how to assert maximum allowed FreezePanes duration in an xUnit test for Aspose.Cells workbooks.

using System;
using System.Diagnostics;
using Aspose.Cells;

// C# sample that creates a workbook, times two FreezePanes overloads with System.Diagnostics.Stopwatch, outputs the elapsed milliseconds to the console, and saves the file as FreezePanesTiming.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Record time for the first FreezePanes call
        Stopwatch stopwatch = Stopwatch.StartNew();
        worksheet.FreezePanes(3, 3, 3, 3);
        stopwatch.Stop();
        Console.WriteLine($"First FreezePanes call duration: {stopwatch.ElapsedMilliseconds} ms");

        // Record time for a second FreezePanes call using cell name
        stopwatch.Restart();
        worksheet.FreezePanes("E5", 5, 5);
        stopwatch.Stop();
        Console.WriteLine($"Second FreezePanes call duration: {stopwatch.ElapsedMilliseconds} ms");

        // Save the workbook
        workbook.Save("FreezePanesTiming.xlsx");
    }
}
