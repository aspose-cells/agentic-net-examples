using System;
using Aspose.Cells;

namespace AsposeCellsOptimizationDemo
{
    class Program
    {
        static void Main()
        {
            // ---------- Load the workbook with formula parsing disabled ----------
            // This skips the expensive formula parsing step during load.
            LoadOptions loadOptions = new LoadOptions
            {
                ParsingFormulaOnOpen = false
            };
            Workbook workbook = new Workbook("input.xlsx", loadOptions);

            // ---------- Optimize calculation settings ----------
            // Disable automatic calculation on open (not needed because we will calculate manually).
            workbook.Settings.FormulaSettings.CalculateOnOpen = false;

            // Enable calculation chain – useful when the same workbook is recalculated many times
            // after only a few cells have changed. The chain is built on first calculation.
            workbook.Settings.FormulaSettings.EnableCalculationChain = true;

            // Create calculation options to fine‑tune the calculation engine.
            CalculationOptions calcOptions = new CalculationOptions
            {
                // Reduce recursion depth if you know formulas are shallow; this can improve speed.
                Recursive = false,

                // Smaller stack size can avoid overhead for deep dependency trees.
                // Adjust according to your workbook complexity.
                CalcStackSize = 100,

                // Ignore errors to prevent interruptions during bulk calculation.
                IgnoreError = true
            };

            // ---------- Perform the calculation ----------
            // This calculates all formulas using the optimized settings.
            workbook.CalculateFormula(calcOptions);

            // ---------- Save the workbook ----------
            // The workbook is saved without triggering another calculation (CalculateOnSave = false).
            workbook.Settings.FormulaSettings.CalculateOnSave = false;
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}