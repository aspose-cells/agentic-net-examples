using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook (adjust the path as needed)
        Workbook workbook = new Workbook("input.xlsx");

        // Clear cached calculation results and perform a fresh calculation
        RecalculateWorkbook(workbook);

        // Save the workbook after the fresh calculation
        workbook.Save("output.xlsx");
    }

    static void RecalculateWorkbook(Workbook workbook)
    {
        // Preserve the original ForceFullCalculation setting
        FormulaSettings formulaSettings = workbook.Settings.FormulaSettings;
        bool originalForceFull = formulaSettings.ForceFullCalculation;

        // Enable force‑full calculation to ignore any cached results
        formulaSettings.ForceFullCalculation = true;

        // Run calculation – all formulas are recomputed from scratch
        workbook.CalculateFormula();

        // Restore the original setting
        formulaSettings.ForceFullCalculation = originalForceFull;
    }
}