// Title: C# – Add SUM(E2:E100) formula and save workbook with Aspose.Cells
// Description: Demonstrates how to create a new Workbook, set the formula "=SUM(E2:E100)" in a cell, evaluate all formulas, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# formula | SUM(E2:E100) Aspose | set cell formula .NET | calculate workbook formulas | save workbook Aspose.Cells
// Common Searches: Aspose.Cells add SUM formula C# | how to set range formula in Aspose.Cells | calculate and save workbook Aspose .NET | C# write SUM(E2:E100) to cell
// Developer Intent: Insert a SUM formula for cells E2‑E100, compute the result, and persist the workbook.
// Use Cases: Generate a total row for financial reports that updates automatically. | Summarize sensor readings or measurement data across many rows. | Create a dynamic summary sheet that recalculates totals on each workbook open.
// AI Prompts: Write C# code with Aspose.Cells to place "=SUM(E2:E100)" in cell F1, calculate the workbook, and save it as an .xlsx file. | Show how to add multiple aggregate formulas (SUM, AVERAGE, COUNT) to different cells using Aspose.Cells and ensure they are evaluated before saving. | Explain how to modify an existing workbook's formula range and recalculate only the affected cells with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaDemo
{
    // Demonstrates how to create a new Workbook, set the formula "=SUM(E2:E100)" in a cell, evaluate all formulas, and save the file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set a formula that sums column E from rows 2 to 100.
            // The result will be placed in cell F1 (you can choose any cell).
            worksheet.Cells["F1"].Formula = "=SUM(E2:E100)";

            // Calculate all formulas in the workbook so the result is materialized.
            workbook.CalculateFormula();

            // Save the workbook (lifecycle: save)
            workbook.Save("SumColumnE.xlsx");
        }
    }
}
