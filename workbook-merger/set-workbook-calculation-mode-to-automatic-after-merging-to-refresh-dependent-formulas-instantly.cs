// Title: Set Automatic Calculation Mode After Merging Cells with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to merge a range, switch the workbook's FormulaSettings.CalculationMode to Automatic, trigger an immediate recalculation, and save the file so dependent formulas update instantly.
// Keywords: Aspose.Cells | C# | .NET | automatic calculation mode | merge cells | formula recalculation | CalcModeType.Automatic | Workbook.CalculateFormula
// Common Searches: Aspose.Cells set calculation mode automatic C# | recalculate formulas after merging cells Aspose.Cells | how to trigger formula calculation programmatically .NET | merge cells and update dependent formulas Aspose
// Developer Intent: Enable instant formula updates after a cell merge by activating automatic calculation.
// Use Cases: Create a report header that spans multiple columns while keeping totals current. | Combine data blocks during workbook generation without manual refresh steps. | Apply automatic calculation in a data‑processing pipeline after structural changes.
// AI Prompts: Generate C# code using Aspose.Cells to merge a range, set CalculationMode to Automatic, and run CalculateFormula. | Explain the steps to ensure formulas recalculate immediately after merging cells in an Aspose.Cells workbook. | Show how to verify that dependent formulas reflect the merged cells' values in a .NET application.

using System;
using Aspose.Cells;

// Demonstrates how to merge a range, switch the workbook's FormulaSettings.CalculationMode to Automatic, trigger an immediate recalculation, and save the file so dependent formulas update instantly.
class SetCalculationModeAfterMerge
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data and a formula that depends on the cell to be merged
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].Formula = "=A1*2";

        // Merge cells A1:B1 (row 0, column 0, 1 row, 2 columns)
        sheet.Cells.Merge(0, 0, 1, 2);

        // Set calculation mode to Automatic so dependent formulas recalculate instantly
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

        // Trigger calculation immediately
        workbook.CalculateFormula();

        // Save the workbook (lifecycle rule)
        workbook.Save("MergedCalcMode.xlsx");
    }
}
