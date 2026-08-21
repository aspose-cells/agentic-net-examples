// Title: Aspose.Cells for .NET: Disable Auto‑Calc, Update Cells, and Manually Recalculate Formulas
// Description: Demonstrates how to set the workbook to manual calculation mode, change several input cells, invoke Workbook.CalculateFormula to refresh dependent formulas, and save the result using C#.
// Keywords: Aspose.Cells manual calculation | C# disable automatic formula calculation | Workbook.CalculateFormula example | batch cell updates performance | formula recalculation .NET | manual calc mode Aspose.Cells | update dependent formulas C#
// Common Searches: Aspose.Cells turn off auto calculation | manual formula calculation C# Aspose.Cells | Workbook.CalculateFormula usage | how to recalculate formulas after cell changes Aspose.Cells
// Developer Intent: Prevent auto‑recalculation, modify data, then trigger formula evaluation on demand.
// Use Cases: Large data imports where auto‑calc would slow processing; disable it, apply all changes, then call CalculateFormula once. | Spreadsheet templates that require user edits before a final total is computed by the application. | Performance‑critical reporting where only specific summary formulas need to be refreshed after batch updates.
// AI Prompts: Generate C# code that sets Aspose.Cells to manual calculation mode, updates multiple cells, runs Workbook.CalculateFormula, and saves the workbook. | Explain best practices for using Workbook.CalculateFormula efficiently after batch modifications in manual mode. | Show how to retrieve updated cell values after manual recalculation and handle potential calculation errors.

using System;
using Aspose.Cells;

namespace AsposeCellsManualCalcDemo
{
    // Demonstrates how to set the workbook to manual calculation mode, change several input cells, invoke Workbook.CalculateFormula to refresh dependent formulas, and save the result using C#.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate initial values
            cells["A1"].PutValue(5);
            cells["A2"].PutValue(10);
            cells["A3"].PutValue(15);

            // Set formulas that depend on the above cells
            cells["B1"].Formula = "=A1*2";      // Expected 10
            cells["B2"].Formula = "=A2*2";      // Expected 20
            cells["B3"].Formula = "=A3*2";      // Expected 30
            cells["C1"].Formula = "=SUM(B1:B3)"; // Expected 60

            // Disable automatic calculation (set mode to Manual)
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

            // Modify several cells after disabling auto‑calc
            cells["A1"].PutValue(7);   // Change from 5 to 7
            cells["A2"].PutValue(14);  // Change from 10 to 14
            cells["A3"].PutValue(21);  // Change from 15 to 21

            // Manually trigger formula calculation
            workbook.CalculateFormula();

            // Output results to console for verification
            Console.WriteLine("After manual calculation:");
            Console.WriteLine($"B1 = {cells["B1"].IntValue}"); // 14
            Console.WriteLine($"B2 = {cells["B2"].IntValue}"); // 28
            Console.WriteLine($"B3 = {cells["B3"].IntValue}"); // 42
            Console.WriteLine($"C1 (SUM) = {cells["C1"].IntValue}"); // 84

            // Save the workbook (lifecycle save)
            workbook.Save("ManualCalcResult.xlsx", SaveFormat.Xlsx);
        }
    }
}
