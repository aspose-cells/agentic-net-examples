// Title: Aspose.Cells .NET: Clear Formula Cache and Force Full Workbook Recalculation
// Description: Loads a workbook, disables cached calculation data by setting FormulaSettings.ForceFullCalculation to true, runs Workbook.CalculateFormula to recompute every formula, resets the flag to keep Excel behavior normal, and saves the updated file. Includes file‑existence checks and robust error handling.
// Keywords: Aspose.Cells | C# | clear formula cache | ForceFullCalculation | Workbook.CalculateFormula | full workbook recalculation | reset formula settings | programmatic Excel calculation | .NET spreadsheet API | remove cached calculations
// Common Searches: Aspose.Cells clear cached calculation results | force full calculation of workbook in C# | reset ForceFullCalculation after CalculateFormula | recalculate all formulas Aspose.Cells .NET | how to discard formula cache Aspose.Cells
// Developer Intent: The developer needs to discard any stored formula results and then recalculate every formula in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Generate a report after external data sources have changed, ensuring all formulas reflect the latest values. | Prepare a workbook for automated testing where a deterministic calculation state is required. | Update a spreadsheet programmatically and guarantee that totals and aggregates are accurate before the file is opened in Excel.
// AI Prompts: Write C# code that clears cached formula results, forces a full recalculation with Aspose.Cells, and saves the workbook. | Explain how the ForceFullCalculation property works and why it should be reset after calling CalculateFormula. | Provide best‑practice error handling for loading a workbook, checking file existence, forcing full calculation, and saving the result with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads a workbook, disables cached calculation data by setting FormulaSettings.ForceFullCalculation to true, runs Workbook.CalculateFormula to recompute every formula, resets the flag to keep Excel behavior normal, and saves the updated file. Includes file‑existence checks and robust error handling.
    public class ClearCacheAndRecalculate
    {
        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Force a full recalculation, ignoring any cached results
                workbook.Settings.FormulaSettings.ForceFullCalculation = true;

                // Recalculate all formulas in the workbook
                workbook.CalculateFormula();

                // Reset the flag so subsequent saves do not force full recalculation in Excel
                workbook.Settings.FormulaSettings.ForceFullCalculation = false;

                // Save the workbook after recalculation
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during processing: {ex.Message}");
            }
        }
    }

    public class Program
    {
        // Entry point required for the console application
        public static void Main(string[] args)
        {
            try
            {
                ClearCacheAndRecalculate.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
