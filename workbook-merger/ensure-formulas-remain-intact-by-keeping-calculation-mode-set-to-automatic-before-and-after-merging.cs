// Title: How to keep formulas intact by switching to Automatic calculation mode before and after merging cells with Aspose.Cells in C#
// AI Prompts: Write C# code using Aspose.Cells that saves the current formula calculation setting, sets it to Automatic, merges a specified cell range, forces formula recalculation, and then restores the original setting. | Show an example of preserving existing Excel formulas when merging cells in a workbook by temporarily enabling automatic formula calculation with Aspose.Cells. | Demonstrate how to merge cells C1:D2 in a C# Aspose.Cells workbook while ensuring the SUM formula remains functional and the original calculation mode is reinstated.
// Common Searches: Aspose.Cells C# merge cells without losing formulas | set calculation mode to Automatic temporarily during cell merge Aspose.Cells .NET | restore original calculation mode after merging cells in Aspose.Cells workbook
// Tags: auto formula calculation during Aspose.Cells merge | keep formulas after cell merge Aspose.Cells | reset formula settings after cell merge Aspose.Cells | C# merge cell range with formula retention Aspose.Cells | temporary formula setting switch Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaMergeDemo
{
    // // Demonstrates creating a workbook, adding data and a SUM formula, storing the original calculation mode, switching to Automatic, merging cells C1:D2, recalculating formulas, restoring the original mode, and saving the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Add sample data
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].PutValue(30);

            // Add a formula that sums the three values
            cells["B1"].Formula = "=SUM(A1:A3)";

            // Store the original calculation mode (in case it is not Automatic)
            CalcModeType originalMode = workbook.Settings.FormulaSettings.CalculationMode;

            // Ensure calculation mode is Automatic before merging (feature rule)
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

            // Perform a merge operation (use Cells.Merge method rule)
            // Merge cells C1:D2 (rows 0‑1, columns 2‑3)
            cells.Merge(0, 2, 2, 2);

            // Optionally, recalculate formulas now that the mode is Automatic
            // (not strictly required because Automatic mode calculates on demand)
            workbook.CalculateFormula();

            // Restore the original calculation mode after merging
            workbook.Settings.FormulaSettings.CalculationMode = originalMode;

            // Save the workbook (save rule)
            workbook.Save("FormulaMergeResult.xlsx");
        }
    }
}
