// Title: Copy a worksheet row and automatically recalculate dependent formulas using Aspose.Cells for .NET
// AI Prompts: Duplicate the first row to a new position and invoke workbook.CalculateFormula() to refresh all dependent formulas in C# with Aspose.Cells. | After copying rows, call RefreshDynamicArrayFormulas(true) to update any spill ranges before saving the workbook programmatically. | Programmatically copy a row, recalculate the sheet, and save the workbook as an .xlsx file using Aspose.Cells in .NET.
// Common Searches: asp.net aspose.cells copy row and recalculate formulas | c# how to refresh dynamic array formulas after row duplication with Aspose.Cells | recalculate workbook after copying rows using Aspose.Cells .NET | copy rows and update dependent formulas in Excel via Aspose.Cells C#
// Tags: copy rows workbook.CalculateFormula Aspose.Cells | refresh dynamic array formulas Aspose.Cells C# | row duplication formula update .NET | worksheet row copy recalc formulas

using System;
using Aspose.Cells;

// The example creates a workbook, fills cells A1‑A3, adds a SUM formula in B1, calculates it, copies the first row to row 5, recalculates all formulas, refreshes any dynamic array formulas, and saves the file as CopyRows_Recalculated.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle create)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some data
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);
        sheet.Cells["A3"].PutValue(30);

        // Add a formula that depends on the above cells
        sheet.Cells["B1"].Formula = "=SUM(A1:A3)";

        // Initial calculation so the formula has a value
        workbook.CalculateFormula();

        // Copy the first row (index 0) to row 5 (index 4)
        // This copies data and formats of the whole row
        sheet.Cells.CopyRows(sheet.Cells, 0, 4, 1);

        // Recalculate after the copy to update any dependent formulas
        workbook.CalculateFormula();

        // Refresh dynamic array formulas if any exist (optional but ensures spill ranges are updated)
        workbook.RefreshDynamicArrayFormulas(true);

        // Save the workbook (lifecycle save)
        workbook.Save("CopyRows_Recalculated.xlsx");
    }
}
