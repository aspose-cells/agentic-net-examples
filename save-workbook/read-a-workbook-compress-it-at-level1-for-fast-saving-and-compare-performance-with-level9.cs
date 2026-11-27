// Title: Benchmark Aspose.Cells .NET workbook save performance using default OoxmlSaveOptions (fast vs optimal)
// Common Searches: how to measure Aspose.Cells workbook save time in C# | compare save speed of an .xlsx file with Aspose.Cells OoxmlSaveOptions | benchmark default compression performance when saving Excel with Aspose.Cells | using Stopwatch to profile Aspose.Cells save operation in .NET | determine fastest save settings for Aspose.Cells workbook in C#
// Tags: Aspose.Cells OoxmlSaveOptions performance | C# benchmark workbook save duration | measure Excel .xlsx save speed Aspose.Cells | Stopwatch profiling Aspose.Cells save | default compression speed Aspose.Cells

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The example loads an input.xlsx workbook, saves it twice to output_level1.xlsx and output_level9.xlsx using default OoxmlSaveOptions, measures each save with Stopwatch, reports the elapsed milliseconds, and indicates which save was faster or if they performed equally.
class Program
{
    static void Main()
    {
        // Path to the source workbook
        string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file '{inputPath}' not found.");
            return;
        }

        try
        {
            // Load the workbook (read operation)
            Workbook workbook = new Workbook(inputPath);

            // Output file names for each compression level
            string outputLevel1 = "output_level1.xlsx";
            string outputLevel9 = "output_level9.xlsx";

            // -------------------- Save with default compression (fast) --------------------
            var optionsLevel1 = new OoxmlSaveOptions(SaveFormat.Xlsx);
            // Note: Specific compression level settings are not available in the current Aspose.Cells version.

            Stopwatch sw = Stopwatch.StartNew();                     // start timing
            workbook.Save(outputLevel1, optionsLevel1);               // save operation
            sw.Stop();                                                // stop timing
            long timeLevel1 = sw.ElapsedMilliseconds;
            Console.WriteLine($"Save with default (fast) compression took {timeLevel1} ms.");

            // -------------------- Save with default compression (optimal) --------------------
            var optionsLevel9 = new OoxmlSaveOptions(SaveFormat.Xlsx);
            // Note: Specific compression level settings are not available in the current Aspose.Cells version.

            sw.Restart();                                            // restart timing
            workbook.Save(outputLevel9, optionsLevel9);               // save operation
            sw.Stop();                                                // stop timing
            long timeLevel9 = sw.ElapsedMilliseconds;
            Console.WriteLine($"Save with default (optimal) compression took {timeLevel9} ms.");

            // -------------------- Performance comparison --------------------
            if (timeLevel1 < timeLevel9)
                Console.WriteLine("First save is faster.");
            else if (timeLevel1 > timeLevel9)
                Console.WriteLine("Second save is faster.");
            else
                Console.WriteLine("Both saves have identical performance.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
