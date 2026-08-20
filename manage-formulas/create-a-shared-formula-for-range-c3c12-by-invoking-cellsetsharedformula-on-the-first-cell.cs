// Title: C# – Set a Shared Formula for C3:C12 Using Cell.SetSharedFormula in Aspose.Cells
// Description: Demonstrates how to create a new workbook, populate B3:B12 with incremental values, apply a shared formula "=B3*2" to the range C3:C12 via Cell.SetSharedFormula, recalculate the sheet, output the results, and save the file as SharedFormulaDemo.xlsx.
// Keywords: Aspose.Cells | Cell.SetSharedFormula | shared formula C3:C12 | C# Aspose.Cells example | set shared formula .NET | calculate formulas Aspose.Cells | Excel shared formula programmatically | performance optimization formulas | GitHub Aspose.Cells sample
// Common Searches: how to use Cell.SetSharedFormula in C# | Aspose.Cells shared formula for a column range | set shared formula C3 to C12 Aspose.Cells | apply one formula to many cells Aspose.Cells .NET | recalculate workbook after setting shared formula
// Developer Intent: Apply a single shared formula to cells C3 through C12 and compute the results with Aspose.Cells for .NET.
// Use Cases: Generate large Excel reports where the same calculation repeats across rows, reducing memory and processing overhead. | Create a template that automatically updates column C based on values entered in column B. | Programmatically build workbooks, apply shared formulas, verify calculations, and export the file.
// AI Prompts: Provide C# code that uses Aspose.Cells to set a shared formula starting at C3 for ten rows and then calculates the workbook. | Explain the meaning of the rowCount and columnCount parameters in Cell.SetSharedFormula and how they define the target range. | Show how to verify the calculated values of a shared formula in Aspose.Cells and save the workbook to disk.

using Aspose.Cells;
using System;

// Demonstrates how to create a new workbook, populate B3:B12 with incremental values, apply a shared formula "=B3*2" to the range C3:C12 via Cell.SetSharedFormula, recalculate the sheet, output the results, and save the file as SharedFormulaDemo.xlsx.
class SetSharedFormulaDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data in column B (B3:B12) that the formula will use
        for (int row = 2; row < 12; row++) // zero‑based index: row 2 = Excel row 3
        {
            cells[row, 1].PutValue(row - 1); // B3=2, B4=3, ..., B12=11
        }

        // Set a shared formula starting at C3 that multiplies the corresponding B cell by 2
        // Parameters: (formula, rowNumber, columnNumber)
        // rowNumber = 10 (C3 to C12), columnNumber = 1 (single column)
        Cell firstCell = cells["C3"];
        firstCell.SetSharedFormula("=B3*2", 10, 1);

        // Calculate all formulas so that values are updated
        workbook.CalculateFormula();

        // Output the results for verification
        for (int row = 2; row < 12; row++)
        {
            Console.WriteLine($"C{row + 1} = {cells[row, 2].Value}");
        }

        // Save the workbook
        workbook.Save("SharedFormulaDemo.xlsx");
    }
}
