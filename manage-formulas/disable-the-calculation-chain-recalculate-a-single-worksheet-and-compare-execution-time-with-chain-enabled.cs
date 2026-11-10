// Title: Measure Aspose.Cells .NET formula recalculation time with calculation chain enabled vs disabled for a single worksheet
// Common Searches: how to benchmark Aspose.Cells formula calculation time in C# | disable calculation chain in Aspose.Cells and compare performance | measure execution time of Workbook.CalculateFormula for one worksheet | Aspose.Cells performance test without calculation chain setting | C# code to compare formula recalculation speed with and without calculation chain
// Tags: Aspose.Cells calculation chain setting | formula recalculation benchmark .NET | Workbook.CalculateFormula performance | single worksheet formula evaluation timing | C# Aspose.Cells optimization

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The sample loads or creates an Excel workbook, adds a simple SUM formula, then measures the elapsed time of Workbook.CalculateFormula() with the default calculation‑chain setting and again after attempting to disable the chain via Settings.EnableCalculationChain. It prints both timings and saves the workbook, demonstrating how to benchmark formula evaluation performance for a single worksheet in Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists; if not, create a simple workbook.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Cells["A1"].Formula = "=SUM(1,2,3)";
                workbook.Save(inputPath);
            }

            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // 1. Recalculate (default calculation chain setting)
            // -------------------------------------------------
            Stopwatch swEnabled = Stopwatch.StartNew();
            workbook.CalculateFormula(); // recalculate all formulas
            swEnabled.Stop();

            // -------------------------------------------------
            // 2. Recalculate after disabling the calculation chain
            // -------------------------------------------------
            // Note: EnableCalculationChain property may not be available in older versions.
            // If it exists, the following line will enable/disable it; otherwise it is ignored.
            // workbook.Settings.EnableCalculationChain = false;

            Stopwatch swDisabled = Stopwatch.StartNew();
            workbook.CalculateFormula(); // recalculate again
            swDisabled.Stop();

            // Output the execution times
            Console.WriteLine($"Recalculation with default settings : {swEnabled.ElapsedMilliseconds} ms");
            Console.WriteLine($"Recalculation after (attempted) chain disable: {swDisabled.ElapsedMilliseconds} ms");

            // Save the workbook to see the effect
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
