// Title: Benchmark the effect of enabling the calculation chain on Aspose.Cells workbook.CalculateFormula performance in C#
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, activates the calculation chain (if the API provides it), runs workbook.CalculateFormula, and logs the elapsed milliseconds using Stopwatch. | Write a C# performance test that calls workbook.CalculateFormula twice, captures the duration of each call, and prints the time difference to the console. | Adapt the sample to compare calculation speed before and after setting WorkbookSettings.EnableCalculationChain = true (when available) and display the resulting speedup.
// Common Searches: how to measure Aspose.Cells CalculateFormula execution time in a .NET application | performance impact of Aspose.Cells calculation chain on Excel formula evaluation | compare first and second workbook.CalculateFormula run times using C# | enable calculation chain Aspose.Cells and benchmark formula calculation speed | C# code to benchmark Aspose.Cells workbook-wide calculation
// Tags: Aspose.Cells calculation chain activation | Aspose.Cells workbook.CalculateFormula timing | C# timing of Excel formula evaluation | Aspose.Cells measurement of calculation speed | Aspose.Cells workbook-wide formula calculation speed

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file with Aspose.Cells, runs workbook.CalculateFormula twice while measuring each execution with Stopwatch, prints the elapsed times and their difference, optionally enables the calculation chain if supported, and saves the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // -------------------------------------------------
            // 1. Measure calculation time with current settings
            // -------------------------------------------------
            Stopwatch stopwatch = Stopwatch.StartNew();

            // Perform workbook-wide calculation
            workbook.CalculateFormula();

            stopwatch.Stop();
            long firstCalcTime = stopwatch.ElapsedMilliseconds;

            // -------------------------------------------------
            // 2. Measure calculation time after a second calculation
            // -------------------------------------------------
            // (Aspose.Cells automatically manages the calculation chain;
            //  the property EnableCalculationChain is not available in this version)
            stopwatch.Restart();

            // Perform workbook-wide calculation again
            workbook.CalculateFormula();

            stopwatch.Stop();
            long secondCalcTime = stopwatch.ElapsedMilliseconds;

            // -------------------------------------------------
            // 3. Output performance results
            // -------------------------------------------------
            Console.WriteLine($"First calculation time: {firstCalcTime} ms");
            Console.WriteLine($"Second calculation time: {secondCalcTime} ms");
            Console.WriteLine($"Performance difference: {firstCalcTime - secondCalcTime} ms");

            // -------------------------------------------------
            // 4. Save the workbook (optional)
            // -------------------------------------------------
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors to prevent the program from crashing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
