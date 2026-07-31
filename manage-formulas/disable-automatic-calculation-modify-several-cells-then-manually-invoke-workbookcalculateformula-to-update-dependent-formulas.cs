// Title: Manual Formula Calculation in Aspose.Cells for .NET – Disable Auto‑Calc, Update Cells, Trigger CalculateFormula
// Description: Shows how to set a workbook to manual calculation mode, change cell values, assign formulas, and explicitly call Workbook.CalculateFormula to evaluate dependent formulas before saving.
// Keywords: Aspose.Cells manual calculation | CalcModeType.Manual .NET | Workbook.CalculateFormula example | disable automatic formula evaluation | prevent calculation on save | C# Aspose.Cells formula performance
// Common Searches: Aspose.Cells turn off automatic calculation | how to recalculate formulas manually in Aspose.Cells | disable calculation on save Aspose.Cells .NET | manual mode CalcModeType example | C# Aspose.Cells calculate after bulk update
// Developer Intent: The developer needs to stop automatic formula evaluation, modify cell data, and then run a single manual calculation to update dependent formulas.
// Use Cases: Bulk‑update many cells without triggering repeated calculations, then compute all formulas once. | Generate a report where formulas must be evaluated only after all input values are populated. | Improve performance in server‑side processing by disabling calculation on save and invoking CalculateFormula explicitly.
// AI Prompts: Provide C# code that sets Aspose.Cells workbook to manual calculation, updates cells, and calls Workbook.CalculateFormula. | Explain when and how to use CalcModeType.Manual and CalculateOnSave = false in Aspose.Cells. | Show how to recalculate formulas for a specific worksheet after editing cell values in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsManualCalculationDemo
{
    // Shows how to set a workbook to manual calculation mode, change cell values, assign formulas, and explicitly call Workbook.CalculateFormula to evaluate dependent formulas before saving.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Disable automatic calculation
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;
            // Optional: prevent calculation on save for performance
            workbook.Settings.FormulaSettings.CalculateOnSave = false;

            // Populate some cells with values
            cells["A1"].PutValue(5);
            cells["A2"].PutValue(10);
            cells["A3"].PutValue(15);

            // Set formulas that depend on the above values
            cells["B1"].Formula = "=A1*2";          // Expected 10
            cells["B2"].Formula = "=A2+5";          // Expected 15
            cells["B3"].Formula = "=SUM(A1:A3)";    // Expected 30

            // At this point formulas are not calculated because calculation mode is Manual

            // Manually calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Verify results (optional console output)
            Console.WriteLine("B1 = " + cells["B1"].Value); // 10
            Console.WriteLine("B2 = " + cells["B2"].Value); // 15
            Console.WriteLine("B3 = " + cells["B3"].Value); // 30

            // Save the workbook
            workbook.Save("ManualCalculationResult.xlsx");
        }
    }
}
