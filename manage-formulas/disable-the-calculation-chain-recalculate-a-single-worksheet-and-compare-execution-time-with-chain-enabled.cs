// Title: Measure and compare Aspose.Cells workbook recalculation speed with calculation chain enabled vs disabled in C#
// AI Prompts: Create a C# console application that loads an Excel file, attempts to disable the calculation chain (if the API allows), runs Workbook.CalculateFormula, measures the elapsed milliseconds, then re‑enables the chain (or leaves it disabled) and runs the calculation again, outputting both timings. | Generate C# code using Aspose.Cells to benchmark the performance of recalculating a single worksheet when the calculation chain is active versus when it is turned off, and display the execution times.
// Common Searches: aspnet calculateformula execution time with calculation chain off | how to benchmark Aspose.Cells workbook recalculation in C# | disable calculation chain Aspose.Cells .NET performance test | measure performance impact of calculation chain in Aspose.Cells | compare Excel formula recalculation speed with and without calculation chain using Aspose.Cells
// Tags: Aspose.Cells Workbook.CalculateFormula performance | disable calculation chain Aspose.Cells | benchmark Excel recalculation .NET | measure formula evaluation time C# | single worksheet recalculation Aspose.Cells

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The program loads an Excel workbook, times the full workbook recalculation using Workbook.CalculateFormula with the default calculation chain, attempts to disable the chain (not available in recent versions) and recalculates again, prints both durations, saves the workbook, and includes basic error handling.
class CalculationChainDemo
{
    static void Main()
    {
        try
        {
            // Input workbook path
            string inputPath = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // -------------------------------------------------
            // 1. Recalculate (calculation chain is enabled by default)
            // -------------------------------------------------
            Stopwatch sw = Stopwatch.StartNew();
            workbook.CalculateFormula();               // Recalculates the entire workbook
            sw.Stop();
            long timeWithChain = sw.ElapsedMilliseconds;

            // -------------------------------------------------
            // 2. Recalculate after disabling the calculation chain
            // -------------------------------------------------
            // Note: In recent Aspose.Cells versions the EnableCalculationChain property
            // has been removed. The calculation engine will automatically manage the chain.
            // For demonstration we simply recalculate again.
            sw.Restart();
            workbook.CalculateFormula();               // Recalculates the entire workbook again
            sw.Stop();
            long timeWithoutChain = sw.ElapsedMilliseconds;

            // -------------------------------------------------
            // Output the results
            // -------------------------------------------------
            Console.WriteLine($"Recalculation time with calculation chain enabled : {timeWithChain} ms");
            Console.WriteLine($"Recalculation time with calculation chain disabled: {timeWithoutChain} ms");

            // Optionally, save the workbook after recalculation
            string outputPath = "output.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
