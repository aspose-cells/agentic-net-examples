// Title: Benchmark Aspose.Cells workbook save speed with OoxmlCompressionLevel 1 (fast) vs Level 9 (maximum) in C#
// AI Prompts: Write a C# console program that loads an Excel file with Aspose.Cells, saves it using OoxmlSaveOptions with CompressionType set to Level1, records the elapsed time with Stopwatch, then repeats the save with CompressionType Level9 and prints both durations. | Create a performance test script for Aspose.Cells that compares the execution time of workbook.Save when using fast (Level1) and maximum (Level9) OOXML compression, outputting the results to the console. | Generate C# code that demonstrates how to configure OoxmlSaveOptions for different compression levels, measure each save operation, and log the timing information for analysis.
// Common Searches: how long does Aspose.Cells take to save an .xlsx with Level1 compression compared to Level9 | C# benchmark Aspose.Cells OoxmlSaveOptions compression performance | measure save time for different OOXML compression levels using Aspose.Cells | compare fast and best compression options in Aspose.Cells .NET
// Tags: Aspose.Cells OoxmlSaveOptions compression performance | C# benchmark workbook save time | OOXML compression level comparison Aspose | measure Aspose.Cells save duration | Level1 vs Level9 Excel compression .NET

using System;
using System.Diagnostics;
using Aspose.Cells;

// The example loads an Excel workbook, saves it twice with OoxmlSaveOptions using CompressionType Level1 and Level9, measures each save with Stopwatch, and writes the elapsed milliseconds for both compression levels to the console.
class CompressionPerformanceDemo
{
    static void Main()
    {
        // Path to the source workbook
        string inputPath = "input.xlsx";

        // Load the workbook from file
        Workbook workbook = new Workbook(inputPath);

        // -------------------- Save with Level1 compression (fast) --------------------
        OoxmlSaveOptions level1Options = new OoxmlSaveOptions();
        level1Options.CompressionType = OoxmlCompressionType.Level1; // fastest, least compression

        Stopwatch swLevel1 = Stopwatch.StartNew();
        workbook.Save("output_Level1.xlsx", level1Options);
        swLevel1.Stop();

        // -------------------- Save with Level9 compression (best) --------------------
        OoxmlSaveOptions level9Options = new OoxmlSaveOptions();
        level9Options.CompressionType = OoxmlCompressionType.Level9; // best compression, slowest

        Stopwatch swLevel9 = Stopwatch.StartNew();
        workbook.Save("output_Level9.xlsx", level9Options);
        swLevel9.Stop();

        // Output the timing results
        Console.WriteLine($"Level1 save time: {swLevel1.ElapsedMilliseconds} ms");
        Console.WriteLine($"Level9 save time: {swLevel9.ElapsedMilliseconds} ms");
    }
}
