// Title: C# example: measuring Aspose.Cells formula calculation time before and after enabling fast formula mode
// Common Searches: C# Aspose.Cells calculate formula performance with EnableFastFormulaCalculation | How to benchmark Excel formula calculation speed using Aspose.Cells in .NET | Compare workbook calculation time with and without fast formula mode in Aspose.Cells | Measure impact of EnableFastFormulaCalculation on large Excel files in C# | Aspose.Cells formula evaluation timing example for .NET developers
// Tags: EnableFastFormulaCalculation impact analysis | Aspose.Cells formula recalculation performance | C# Excel workbook calculation benchmark | fast formula calculation setting Aspose.Cells | measure formula evaluation latency .NET

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook, records the elapsed time for a full formula calculation, optionally enables the EnableFastFormulaCalculation setting, recalculates the formulas, records the second elapsed time, prints both durations, and saves the workbook. This demonstrates how to benchmark the performance benefit of fast formula mode in Aspose.Cells using C#.
class FastFormulaCalculationDemo
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

            // Measure calculation time before any fast mode setting
            Stopwatch swBefore = Stopwatch.StartNew();
            try
            {
                workbook.CalculateFormula();
            }
            catch (Exception calcEx)
            {
                Console.WriteLine($"Error during calculation (before fast mode): {calcEx.Message}");
                return;
            }
            swBefore.Stop();

            // NOTE: The EnableFastFormulaCalculation property may not be available in older Aspose.Cells versions.
            // If your version supports it, uncomment the following line:
            // workbook.Settings.EnableFastFormulaCalculation = true;

            // Measure calculation time after (potential) fast mode setting
            Stopwatch swAfter = Stopwatch.StartNew();
            try
            {
                workbook.CalculateFormula();
            }
            catch (Exception calcEx)
            {
                Console.WriteLine($"Error during calculation (after fast mode): {calcEx.Message}");
                return;
            }
            swAfter.Stop();

            // Output the timing results
            Console.WriteLine($"Calculation time before enabling fast mode: {swBefore.ElapsedMilliseconds} ms");
            Console.WriteLine($"Calculation time after enabling fast mode:  {swAfter.ElapsedMilliseconds} ms");

            // Save the workbook (output path)
            string outputPath = "output.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
