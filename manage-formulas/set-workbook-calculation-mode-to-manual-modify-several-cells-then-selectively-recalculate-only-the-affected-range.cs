using System;
using Aspose.Cells;

namespace AsposeCellsManualCalcDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // -------------------------------------------------
            // Configure formula settings
            // -------------------------------------------------
            // Set calculation mode to Manual
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

            // Enable calculation chain to allow selective recalculation
            workbook.Settings.FormulaSettings.EnableCalculationChain = true;

            // Optional: prevent automatic calculation on save
            workbook.Settings.FormulaSettings.CalculateOnSave = false;

            // -------------------------------------------------
            // Add initial data and formulas
            // -------------------------------------------------
            cells["A1"].PutValue(10);               // Input value
            cells["A2"].PutValue(20);               // Input value
            cells["B1"].Formula = "=A1*2";          // Dependent formula
            cells["B2"].Formula = "=A2*3";          // Dependent formula
            cells["C1"].Formula = "=B1+B2";         // Formula depending on B1 and B2

            // Initial calculation (required to build the calculation chain)
            workbook.CalculateFormula();

            // -------------------------------------------------
            // Modify a subset of cells
            // -------------------------------------------------
            cells["A1"].PutValue(15);   // Change only A1; only B1 and C1 should be recalculated
            cells["A2"].PutValue(25);   // Change only A2; only B2 and C1 should be recalculated

            // -------------------------------------------------
            // Selectively recalculate affected formulas
            // -------------------------------------------------
            // Because EnableCalculationChain is true, this call recalculates only the cells
            // whose values depend on the modified cells.
            workbook.CalculateFormula();

            // -------------------------------------------------
            // Save the workbook (lifecycle rule: save)
            // -------------------------------------------------
            workbook.Save("ManualCalcSelectiveRecalc.xlsx", SaveFormat.Xlsx);
        }
    }
}