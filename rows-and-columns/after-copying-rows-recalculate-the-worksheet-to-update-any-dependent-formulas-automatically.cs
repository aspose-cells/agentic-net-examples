// Title: C# – Copy Rows and Recalculate Formulas with Aspose.Cells for .NET
// Description: Shows how to copy a block of rows in an Aspose.Cells workbook, then automatically update all dependent formulas—including dynamic array spills—by calling Workbook.CalculateFormula and Workbook.RefreshDynamicArrayFormulas before saving the file.
// Keywords: Aspose.Cells | CopyRows | CalculateFormula | RefreshDynamicArrayFormulas | .NET | C# | update formulas after row copy | recalculate worksheet | dynamic array formulas | Excel automation
// Common Searches: Aspose.Cells copy rows and recalculate formulas | Workbook.CalculateFormula after CopyRows | RefreshDynamicArrayFormulas usage example | How to update formulas after copying rows in C# | CopyRows method Aspose.Cells tutorial
// Developer Intent: Automatically refresh all formulas after rows are copied so that totals and references reflect the new data layout.
// Use Cases: Duplicate a table with summary formulas and ensure the SUM, AVERAGE, etc., recalculate correctly. | Copy a template section that contains dynamic array formulas and refresh spill ranges before generating a report. | Programmatically shift financial data rows and update dependent calculations for an up‑to‑date snapshot.
// AI Prompts: Generate C# code that copies rows in an Aspose.Cells worksheet and then recalculates all formulas, including dynamic arrays. | Explain the difference between Workbook.CalculateFormula and Workbook.RefreshDynamicArrayFormulas after a CopyRows operation. | Provide a method to recalculate only the cells affected by a row copy using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Shows how to copy a block of rows in an Aspose.Cells workbook, then automatically update all dependent formulas—including dynamic array spills—by calling Workbook.CalculateFormula and Workbook.RefreshDynamicArrayFormulas before saving the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Populate some sample data
        ws.Cells["A1"].PutValue(10);
        ws.Cells["A2"].PutValue(20);
        ws.Cells["A3"].PutValue(30);

        // Add a formula that depends on the above data
        ws.Cells["B1"].Formula = "=SUM(A1:A3)";

        // Copy the first three rows (0‑based index) to rows starting at index 3 (row 4)
        // Parameters: source cells, source start row, destination start row, number of rows to copy
        ws.Cells.CopyRows(ws.Cells, 0, 3, 3);

        // Recalculate all formulas in the workbook after the copy operation
        wb.CalculateFormula();

        // Refresh dynamic array formulas if any exist (optional but ensures spill ranges are updated)
        wb.RefreshDynamicArrayFormulas(true);

        // Save the resulting workbook
        wb.Save("CopyRowsAndRecalc.xlsx");
    }
}
