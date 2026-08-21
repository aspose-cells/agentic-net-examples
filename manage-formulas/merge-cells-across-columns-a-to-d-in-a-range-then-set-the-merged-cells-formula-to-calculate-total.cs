// Title: C# Aspose.Cells Example: Merge A1:D1 and Apply a SUM Formula
// Description: Demonstrates how to create a workbook with Aspose.Cells for .NET, fill cells A1‑D1, merge them into a single range, assign a =SUM(A1:D1) formula to the merged cell, calculate the result, output it to the console, and save the file as MergedCellsWithTotal.xlsx.
// Keywords: Aspose.Cells | C# | .NET | merge cells | cells.Merge | SUM formula | calculate total | Excel automation | worksheet formula | CalculateFormula | save workbook
// Common Searches: Aspose.Cells merge cells and set formula C# | How to apply SUM to merged cells using Aspose.Cells | C# code to merge A1:D1 and calculate total | Aspose.Cells example for merged range formula | Calculate sum after merging cells in .NET
// Developer Intent: Merge a specific cell range and attach a SUM formula that computes the total of the original cells.
// Use Cases: Create a multi‑column header that automatically shows the sum of its underlying data. | Generate a summary row where the label spans several columns and the total is calculated in the merged cell. | Build financial or inventory reports that need a merged label with a dynamic aggregate value.
// AI Prompts: Generate C# code using Aspose.Cells to merge cells A1:D1, set a dynamic =SUM(A1:D1) formula, and recalculate when source values change. | Provide an Aspose.Cells for .NET snippet that merges a range, applies a formula, handles possible errors, and saves the workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsMergeAndFormulaDemo
{
    // Demonstrates how to create a workbook with Aspose.Cells for .NET, fill cells A1‑D1, merge them into a single range, assign a =SUM(A1:D1) formula to the merged cell, calculate the result, output it to the console, and save the file as MergedCellsWithTotal.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some sample values in the range A1:D1
            cells["A1"].PutValue(10);
            cells["B1"].PutValue(20);
            cells["C1"].PutValue(30);
            cells["D1"].PutValue(40);

            // Merge cells from column A (0) to column D (3) in row 0 (first row)
            // Parameters: firstRow, firstColumn, totalRows, totalColumns
            cells.Merge(firstRow: 0, firstColumn: 0, totalRows: 1, totalColumns: 4);

            // After merging, the merged cell is referenced by the upper‑left cell (A1)
            // Set a formula in the merged cell to calculate the total of the original range
            cells["A1"].Formula = "=SUM(A1:D1)";

            // Calculate formulas so the result is stored in the cell
            workbook.CalculateFormula();

            // Optional: display the calculated total in the console
            Console.WriteLine("Total of merged cells: " + cells["A1"].Value);

            // Save the workbook to a file
            workbook.Save("MergedCellsWithTotal.xlsx");
        }
    }
}
