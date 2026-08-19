// Title: Keep formulas intact when merging cells – set CalculationMode to Automatic in Aspose.Cells for .NET
// Description: Shows how to preserve existing formulas while merging cells by temporarily switching the workbook’s CalculationMode to Automatic, then restoring the original setting, using Aspose.Cells for .NET.
// Keywords: Aspose.Cells merge cells | C# calculation mode automatic | preserve formulas Aspose.Cells | cell merge formula integrity | Aspose.Cells workbook settings | C# Excel automation | Aspose.Cells calculation mode | merge cells without breaking formulas | Aspose.Cells .NET example | Excel formula recalculation
// Common Searches: Aspose.Cells merge cells keep formula | Set CalculationMode Automatic before merging Aspose.Cells | Restore original calculation mode after merge C# | How to prevent formula loss when merging cells Aspose.Cells | C# Aspose.Cells merge header cells formula | Aspose.Cells calculation mode best practice
// Developer Intent: Ensure that merging cells does not disrupt existing formulas by temporarily enabling automatic calculation and then returning to the workbook’s original calculation setting.
// Use Cases: Create a report header that spans multiple columns while a SUM formula referencing those cells stays accurate. | Adjust the layout of a generated financial model without affecting dependent formulas. | Apply custom workbook settings, perform structural changes, and guarantee the original calculation preferences are preserved.
// AI Prompts: Generate C# code using Aspose.Cells that merges a range of cells and temporarily sets CalculationMode to Automatic to keep formulas working. | Explain why switching to Automatic calculation before a merge prevents formula errors in Aspose.Cells. | Show how to save and restore the original CalculationMode around a merge operation in a .NET spreadsheet automation script.

using System;
using Aspose.Cells;

// Shows how to preserve existing formulas while merging cells by temporarily switching the workbook’s CalculationMode to Automatic, then restoring the original setting, using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle create)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Add sample data and a formula
        cells["A1"].PutValue(10);
        cells["A2"].PutValue(20);
        cells["B1"].Formula = "=SUM(A1:A2)";

        // Preserve the original calculation mode
        CalcModeType originalMode = workbook.Settings.FormulaSettings.CalculationMode;

        // Set calculation mode to Automatic before merging
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

        // Merge cells A1:B1 (row 0, column 0, 1 row, 2 columns)
        cells.Merge(0, 0, 1, 2);

        // Restore the original calculation mode after merging
        workbook.Settings.FormulaSettings.CalculationMode = originalMode;

        // Save the workbook (lifecycle save)
        workbook.Save("MergedWithFormula.xlsx");
    }
}
