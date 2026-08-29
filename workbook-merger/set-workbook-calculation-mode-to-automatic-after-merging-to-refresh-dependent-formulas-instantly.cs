// Title: Set Aspose.Cells workbook to automatic calculation after merging cells to instantly refresh formulas (C#)
// AI Prompts: Combine a range of cells, switch the workbook's formula calculation mode to Automatic, and trigger a full formula recomputation using Aspose.Cells. | In C#, after merging cells with Aspose.Cells, programmatically enable automatic recalculation and force an immediate evaluation of all formulas.
// Common Searches: Aspose.Cells set calculation mode to automatic after cell merge C# | Refresh formulas instantly after merging cells in Aspose.Cells workbook | C# Aspose.Cells automatic recalculation after merging A1:A2 | How to trigger formula recomputation after merging cells with Aspose.Cells | Enable automatic formula evaluation in Aspose.Cells after merging cells
// Tags: cell range merging with automatic formula evaluation Aspose.Cells | Workbook.Settings.FormulaSettings.CalculationMode property | Workbook.CalculateFormula method after merge | Aspose.Cells automatic formula refresh C# | configure Automatic calculation mode in Aspose.Cells workbook

using System;
using Aspose.Cells;

// The example creates a workbook, adds numeric values, defines a SUM formula, merges cells A1:A2, switches the calculation mode to Automatic, forces an immediate formula recomputation, and saves the file as MergedAutomaticCalc.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);

        // Formula that depends on the cells we will merge
        sheet.Cells["B1"].Formula = "=SUM(A1:A2)";

        // Merge the cells A1:A2
        // Parameters: startRow, startColumn, totalRows, totalColumns
        sheet.Cells.Merge(0, 0, 2, 1); // Merges A1 and A2

        // Set calculation mode to Automatic after merging
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

        // Refresh dependent formulas instantly
        workbook.CalculateFormula();

        // Save the workbook (lifecycle: save)
        workbook.Save("MergedAutomaticCalc.xlsx");
    }
}
