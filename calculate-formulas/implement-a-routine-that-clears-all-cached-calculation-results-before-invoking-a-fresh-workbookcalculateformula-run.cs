// Title: Clear Formula Cache and Force Full Recalculation with Aspose.Cells for .NET
// Description: Load a workbook, enable ForceFullCalculation to discard cached results, run Workbook.CalculateFormula for a fresh evaluation, reset the flag, and save the updated file.
// Keywords: Aspose.Cells | C# | .NET | clear formula cache | ForceFullCalculation | full workbook recalculation | CalculateFormula fresh run | reset cached calculations
// Common Searches: Aspose.Cells clear cached calculations | ForceFullCalculation example C# | recalculate all formulas after data change Aspose.Cells | how to reset formula cache in .NET workbook | full calculation mode Aspose.Cells
// Developer Intent: Discard any stored formula results and trigger a complete recomputation of all workbook formulas.
// Use Cases: Refresh formulas after bulk data updates to ensure accurate totals. | Eliminate stale values in automated reporting pipelines. | Guarantee correct calculations when workbook is reused across multiple processing runs.
// AI Prompts: Show me C# code that clears the formula cache and forces a full recalculation using Aspose.Cells. | Explain when to use ForceFullCalculation and its performance impact in Aspose.Cells. | Provide a step‑by‑step guide to toggle ForceFullCalculation, run CalculateFormula, and restore default settings.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Load a workbook, enable ForceFullCalculation to discard cached results, run Workbook.CalculateFormula for a fresh evaluation, reset the flag, and save the updated file.
    public class ClearCacheAndRecalculate
    {
        /// <param name="inputPath">Path to the source workbook.</param>
        /// <param name="outputPath">Path where the recalculated workbook will be saved.</param>
        public static void Run(string inputPath, string outputPath)
        {
            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputPath);

            // Enable full calculation to discard any cached results.
            // This forces all formulas to be recomputed on the next CalculateFormula call.
            workbook.Settings.FormulaSettings.ForceFullCalculation = true;

            // Perform the calculation (fresh run)
            workbook.CalculateFormula();

            // Optionally reset the flag if further operations should use default behavior.
            workbook.Settings.FormulaSettings.ForceFullCalculation = false;

            // Save the workbook (lifecycle rule: save)
            workbook.Save(outputPath);
        }

        // Example usage
        public static void Main()
        {
            string sourceFile = "input.xlsx";
            string resultFile = "output.xlsx";

            Run(sourceFile, resultFile);

            Console.WriteLine("Workbook recalculated and saved to: " + resultFile);
        }
    }
}
