// Title: Clear cached formula values and force a full recalculation of an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# method that sets Workbook.Settings.FormulaSettings.ForceFullCalculation to true, runs CalculateFormula, then restores the flag using Aspose.Cells. | Show how to invalidate the internal formula cache before saving an Excel workbook with Aspose.Cells in .NET. | Generate example code that loads an .xlsx file, forces a complete recomputation of all formulas, and writes the updated workbook back.
// Common Searches: Aspose.Cells .NET force full workbook calculation ignoring cached values | C# clear cached formula results before CalculateFormula with Aspose.Cells | How to recalculate all formulas from scratch in an Excel file using Aspose.Cells
// Tags: enable ForceFullCalculation Aspose.Cells | clear cached formula values C# | Aspose.Cells fresh CalculateFormula execution | reset FormulaSettings after full calculation | recompute all formulas from scratch .NET

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example loads an Excel workbook, enables ForceFullCalculation to discard any cached formula results, executes CalculateFormula to recompute all formulas, optionally resets the flag, and saves the workbook with the newly calculated values.
    public class ClearCacheAndRecalculate
    {
        /// <param name="inputPath">Path to the source workbook.</param>
        /// <param name="outputPath">Path where the recalculated workbook will be saved.</param>
        public static void Run(string inputPath, string outputPath)
        {
            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputPath);

            // Ensure that all formulas are recalculated from scratch.
            // Setting ForceFullCalculation to true forces a full calculation
            // and ignores any previously cached values.
            workbook.Settings.FormulaSettings.ForceFullCalculation = true;

            // Perform the calculation (lifecycle rule: calculate)
            workbook.CalculateFormula();

            // Optional: reset the flag if you don't want the workbook to retain this setting.
            workbook.Settings.FormulaSettings.ForceFullCalculation = false;

            // Save the workbook (lifecycle rule: save)
            workbook.Save(outputPath);
        }

        // Example usage
        public static void Main()
        {
            string inputFile = "input.xlsx";
            string outputFile = "output_recalculated.xlsx";

            Run(inputFile, outputFile);

            Console.WriteLine("Workbook recalculated and saved to: " + outputFile);
        }
    }
}
