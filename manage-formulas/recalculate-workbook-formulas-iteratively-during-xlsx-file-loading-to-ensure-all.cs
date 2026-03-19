using System;
using Aspose.Cells;

namespace AsposeCellsIterativeRecalc
{
    class Program
    {
        static void Main()
        {
            // Load the workbook with formula parsing disabled on open.
            // This ensures formulas are not automatically parsed during loading.
            LoadOptions loadOptions = new LoadOptions
            {
                ParsingFormulaOnOpen = false
            };
            Workbook workbook = new Workbook("input.xlsx", loadOptions);

            // Parse all formulas now that the workbook is loaded.
            // ignoreError = false – throw if any formula is invalid.
            workbook.ParseFormulas(false);

            // Enable iterative calculation to resolve circular references
            // and ensure dependent cells are updated correctly.
            FormulaSettings fSettings = workbook.Settings.FormulaSettings;
            fSettings.EnableIterativeCalculation = true;
            fSettings.MaxIteration = 100;   // maximum number of iterations
            fSettings.MaxChange = 0.001;    // convergence threshold

            // Perform full calculation of all formulas.
            // This will iteratively recalculate dependent cells.
            workbook.CalculateFormula();

            // Save the updated workbook.
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}