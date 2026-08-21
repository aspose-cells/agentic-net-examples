// Title: Aspose.Cells C# – Set Workbook Formula Calculation Mode to Manual
// Description: Demonstrates how to switch a workbook’s FormulaSettings.CalculationMode to CalcModeType.Manual, prevent automatic formula evaluation, and trigger calculation explicitly with workbook.CalculateFormula() before saving.
// Keywords: Aspose.Cells | C# manual calculation mode | CalcModeType.Manual | disable automatic formula evaluation | explicit formula calculation | performance optimization | large dataset recalculation
// Common Searches: Aspose.Cells set manual calculation mode | how to disable automatic formula calculation in .NET | manual vs automatic calculation Aspose.Cells | calculate formulas after manual mode
// Developer Intent: Configure a workbook to use manual formula calculation so formulas are evaluated only when the developer calls CalculateFormula.
// Use Cases: Generate a report with thousands of formulas, update data in bulk, then calculate once to reduce processing time. | Prevent intermediate recalculations while programmatically modifying cell values. | Allow user‑driven changes in a UI, then recalculate selectively before exporting the workbook.
// AI Prompts: Write C# code that toggles between Manual and Automatic calculation modes using Aspose.Cells. | Show how to recalculate a specific cell range after setting the workbook to manual mode. | Explain best practices for using manual calculation mode when processing large spreadsheets with Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to switch a workbook’s FormulaSettings.CalculationMode to CalcModeType.Manual, prevent automatic formula evaluation, and trigger calculation explicitly with workbook.CalculateFormula() before saving.
class SetManualCalculationMode
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data and a formula
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);
        sheet.Cells["A3"].Formula = "=A1+A2";

        // Set the calculation mode to Manual
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

        // At this point the formula has not been calculated automatically
        Console.WriteLine("Before manual calculation, A3 value: " + sheet.Cells["A3"].Value);

        // Perform calculation explicitly when needed
        workbook.CalculateFormula();

        Console.WriteLine("After manual calculation, A3 value: " + sheet.Cells["A3"].IntValue);

        // Save the workbook
        workbook.Save("ManualCalculation.xlsx");
    }
}
