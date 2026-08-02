// Title: C# – Copy rows and auto‑recalculate formulas with Aspose.Cells
// Description: Create a workbook, fill cells A1‑A3, add a SUM formula, copy rows 1‑3 to row 5 using Cells.CopyRows, then call Workbook.CalculateFormula to refresh dependent calculations and save as CopyRowsAndRecalc.xlsx.
// Keywords: Aspose.Cells | C# | CopyRows | CalculateFormula | recalculate formulas | duplicate rows | worksheet copy | Excel automation
// Common Searches: Aspose.Cells copy rows C# | recalculate formulas after copying rows Aspose | Workbook.CalculateFormula example | how to duplicate rows with formulas in .NET | copy rows and update sums Aspose.Cells
// Developer Intent: Duplicate a range of rows and immediately recalculate all formulas so that totals and references reflect the new data.
// Use Cases: Copy a data block with a total row and have the total update automatically. | Replicate a template section within a sheet and refresh its summary calculations. | Generate monthly reports by copying previous month’s rows and recalculating all formulas before export.
// AI Prompts: Write C# code that copies rows in an Aspose.Cells worksheet and then runs CalculateFormula to refresh formulas. | Explain the effect of Workbook.CalculateFormula after using Cells.CopyRows in Aspose.Cells. | Show a step‑by‑step example of copying rows with dependent formulas and ensuring they recalculate correctly in .NET.

using System;
using Aspose.Cells;

// Create a workbook, fill cells A1‑A3, add a SUM formula, copy rows 1‑3 to row 5 using Cells.CopyRows, then call Workbook.CalculateFormula to refresh dependent calculations and save as CopyRowsAndRecalc.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some data in the first three rows
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);
        sheet.Cells["A3"].PutValue(30);

        // Add a formula that depends on the data above
        sheet.Cells["B1"].Formula = "=SUM(A1:A3)";

        // Copy rows 0‑2 (A1:B3) to destination starting at row index 4 (row 5 in Excel)
        // (lifecycle rule: copy rows)
        sheet.Cells.CopyRows(sheet.Cells, 0, 4, 3);

        // Recalculate all formulas in the workbook after the copy operation
        // (lifecycle rule: calculate formulas)
        workbook.CalculateFormula();

        // Save the workbook (lifecycle rule: save)
        workbook.Save("CopyRowsAndRecalc.xlsx");
    }
}
