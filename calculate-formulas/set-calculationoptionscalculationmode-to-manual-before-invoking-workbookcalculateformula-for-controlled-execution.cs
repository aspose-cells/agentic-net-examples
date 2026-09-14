// Title: How to set Aspose.Cells workbook to Manual calculation mode and evaluate formulas with CalculateFormula in C#
// AI Prompts: Demonstrate setting workbook.Settings.FormulaSettings.CalculationMode to CalcModeType.Manual, adding a formula, and invoking workbook.CalculateFormula to compute results in C#. | Provide a C# snippet that turns off auto‑recalculation of formulas in Aspose.Cells, inserts a custom formula, and triggers a manual calculation. | Show how to persist manual calculation mode when saving an Aspose.Cells workbook after performing a controlled formula evaluation.
// Common Searches: Aspose.Cells C# set calculation mode to manual before calling CalculateFormula | disable automatic formula recalculation in Aspose.Cells .NET example | manual formula evaluation with Aspose.Cells workbook.CalculateFormula | how to control when formulas are calculated in Aspose.Cells using CalcModeType.Manual | save Aspose.Cells workbook after manual calculation mode is applied
// Tags: Aspose.Cells manual calculation mode | prevent auto formula calculation Aspose.Cells | Workbook.CalculateFormula manual trigger | CalcModeType.Manual C# example | controlled formula calculation Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsCalculationModeDemo
{
    // Creates an in‑memory workbook, sets the calculation mode to Manual, adds a simple addition formula, shows that the cell remains unevaluated until workbook.CalculateFormula() is called, then outputs the result and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (in-memory)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data
            cells["A1"].PutValue(5);
            cells["A2"].PutValue(10);
            cells["B1"].Formula = "=A1+A2"; // Simple addition formula

            // Set the calculation mode to Manual.
            // This ensures that formulas are not calculated automatically
            // and we have full control over when calculation occurs.
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

            // At this point, the formula result is not yet calculated.
            Console.WriteLine("Before manual calculation, B1 value: " + cells["B1"].Value);

            // Perform controlled calculation of all formulas in the workbook.
            workbook.CalculateFormula();

            // After calculation, the formula result is available.
            Console.WriteLine("After manual calculation, B1 value: " + cells["B1"].IntValue);

            // Save the workbook to verify that the calculation mode is persisted.
            workbook.Save("ManualCalculationModeDemo.xlsx");
        }
    }
}
