// Title: Aspose.Cells C# – Verify Manual Calculation Mode Blocks Automatic Formula Evaluation
// Description: Demonstrates how to set CalcModeType.Manual in an Aspose.Cells workbook, confirm that a formula (C1 = A1+B1) remains unevaluated until workbook.CalculateFormula is invoked, and then save the file.
// Keywords: Aspose.Cells manual calculation | CalcModeType.Manual | C# workbook.CalculateFormula | prevent automatic formula evaluation | Aspose.Cells formula settings | manual calc mode test
// Common Searches: Aspose.Cells set manual calculation mode .NET | formula returns null before CalculateFormula | how to trigger calculation manually in Aspose.Cells | verify manual calc mode persists after save
// Developer Intent: Confirm that formulas are not calculated automatically when the workbook is in manual mode and are evaluated only after an explicit CalculateFormula call.
// Use Cases: Unit test to ensure formulas stay unevaluated until manual calculation is requested. | Performance‑critical spreadsheets where calculations are performed on demand. | Saving a workbook with manual mode so that downstream users control when formulas recalculate.
// AI Prompts: Create an MSTest that asserts cells["C1"].Value is null before CalculateFormula and equals 15 after the call. | Generate C# code to switch a workbook from automatic to manual calculation mode and then recalculate a specific range. | Explain how Aspose.Cells stores the cached result of a formula cell when CalcModeType.Manual is active.

using System;
using Aspose.Cells;

// Demonstrates how to set CalcModeType.Manual in an Aspose.Cells workbook, confirm that a formula (C1 = A1+B1) remains unevaluated until workbook.CalculateFormula is invoked, and then save the file.
class ManualCalcModeTest
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Add sample data
        cells["A1"].PutValue(5);
        cells["B1"].PutValue(10);

        // Set a formula that depends on the above cells
        cells["C1"].Formula = "=A1+B1";

        // Set calculation mode to Manual – formulas will not be evaluated automatically
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

        // Verify that the formula cell has no calculated value yet
        Console.WriteLine("Before calling CalculateFormula:");
        Console.WriteLine($"C1 value: {(cells["C1"].Value ?? "null")}");

        // Manually trigger formula calculation
        workbook.CalculateFormula();

        // Verify that the formula has now been evaluated
        Console.WriteLine("After calling CalculateFormula:");
        Console.WriteLine($"C1 value: {cells["C1"].Value}");

        // Save the workbook (optional, demonstrates that the mode is persisted)
        workbook.Save("ManualCalcModeTest.xlsx");
    }
}
