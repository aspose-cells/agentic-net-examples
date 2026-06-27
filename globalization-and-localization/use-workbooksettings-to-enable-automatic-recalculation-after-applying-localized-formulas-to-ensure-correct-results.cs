using System;
using Aspose.Cells;

namespace AsposeCellsLocalizedFormulaDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set the workbook region to German (uses comma as decimal separator, etc.)
            workbook.Settings.Region = CountryCode.Germany;

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Set a value in B1 that will be used by the formula
            cells["B1"].PutValue(10);

            // Set a formula in local (German) format using FormulaLocal.
            // In German the SUM function is "SUMME".
            cells["A1"].FormulaLocal = "=SUMME(B1;5)";

            // Ensure automatic recalculation is enabled.
            // This forces the engine to recalculate formulas whenever needed.
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

            // Recalculate all formulas in the workbook.
            workbook.CalculateFormula();

            // Output the calculated result of the localized formula.
            Console.WriteLine("Result of localized formula in A1: " + cells["A1"].Value);

            // Save the workbook (optional, demonstrates that settings are persisted)
            workbook.Save("LocalizedFormulaResult.xlsx");
        }
    }
}