// Title: Aspose.Cells .NET – Enable Automatic Calculation & Force Full Recalculation
// Description: Demonstrates how to set a workbook to Automatic calculation mode, activate ForceFullCalculation, run an initial CalculateFormula, modify cell data, trigger another full recalculation, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells automatic calculation | ForceFullCalculation .NET | CalcModeType Automatic | Workbook.CalculateFormula example | C# Excel formula recalculation | Aspose.Cells formula settings
// Common Searches: enable automatic formula calculation Aspose.Cells C# | force full workbook recalculation each time Aspose.Cells | recalculate formulas after cell update Aspose.Cells .NET | set calculation mode to automatic Aspose.Cells
// Developer Intent: Configure a workbook to recalculate all formulas automatically and force a complete recompute after every data change.
// Use Cases: Keep financial dashboards up‑to‑date when source values change. | Ensure consistent results in complex models that use volatile functions. | Maintain data integrity in generated reports that span multiple worksheets.
// AI Prompts: Show how to switch between Automatic and Manual calculation modes in Aspose.Cells and trigger a full recalculation. | Provide C# code to recalculate only a specific range after updating cells with Aspose.Cells. | Explain the performance trade‑offs of ForceFullCalculation and suggest when to use it in large workbooks.

using System;
using Aspose.Cells;

// Demonstrates how to set a workbook to Automatic calculation mode, activate ForceFullCalculation, run an initial CalculateFormula, modify cell data, trigger another full recalculation, and save the file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add initial data and a formula
        worksheet.Cells["A1"].PutValue(10);
        worksheet.Cells["A2"].PutValue(20);
        worksheet.Cells["A3"].Formula = "=SUM(A1:A2)";

        // Enable automatic calculation mode (Excel setting)
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

        // Force a full recalculation each time a calculation is triggered
        workbook.Settings.FormulaSettings.ForceFullCalculation = true;

        // Perform the initial full calculation
        workbook.CalculateFormula();

        // Modify data to demonstrate that a full recalculation occurs again
        worksheet.Cells["A1"].PutValue(30);

        // Recalculate after data change
        workbook.CalculateFormula();

        // Save the workbook (lifecycle rule: save)
        workbook.Save("AutomaticFullCalc.xlsx");
    }
}
