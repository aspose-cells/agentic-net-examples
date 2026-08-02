// Title: Set a Shared Formula for C3:C12 Using Cell.SetSharedFormula in Aspose.Cells for .NET
// Description: This C# example creates a workbook, fills B3:B12 with sequential numbers, applies a shared formula "=B3*2" to the range C3:C12 via Cell.SetSharedFormula, recalculates the sheet, prints the results, and saves the file as SharedFormulaDemo.xlsx.
// Keywords: Aspose.Cells SetSharedFormula C# | shared formula .NET example | Cell.SetSharedFormula usage | apply formula to range Aspose | calculate formulas Aspose.Cells
// Common Searches: how to use Cell.SetSharedFormula in Aspose.Cells | shared formula for a column range C# Aspose | apply same formula to multiple cells Aspose.Cells | recalculate workbook after setting shared formula
// Developer Intent: Apply a single formula to the first cell of a range so Aspose.Cells propagates it automatically to all cells in that range.
// Use Cases: Generate column C values by multiplying each corresponding B cell by 2 without looping through each row. | Reduce memory and processing overhead when assigning identical formulas to large ranges. | Ensure calculated results are stored before exporting the workbook.
// AI Prompts: Write C# code that sets a shared formula for D5:D20 referencing column C and saves the workbook with Aspose.Cells. | Show how to modify a shared formula to use absolute references like $B$3 and recalculate the sheet. | Provide an example that retrieves and prints the results of a shared formula after calling workbook.CalculateFormula.

using System;
using Aspose.Cells;

namespace AsposeCellsSharedFormulaDemo
{
    // This C# example creates a workbook, fills B3:B12 with sequential numbers, applies a shared formula "=B3*2" to the range C3:C12 via Cell.SetSharedFormula, recalculates the sheet, prints the results, and saves the file as SharedFormulaDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data in column B (B3:B12) that the shared formula will use
            for (int i = 2; i < 12; i++) // rows are zero‑based; row 2 corresponds to Excel row 3
            {
                cells[i, 1].PutValue(i - 1); // B3 = 2, B4 = 3, ... B12 = 11
            }

            // Get the first cell of the target range (C3)
            Cell firstCell = cells["C3"];

            // Set a shared formula for the range C3:C12 (10 rows, 1 column)
            // The formula references the cell in column B of the same row.
            // Aspose.Cells will adjust the reference for each row automatically.
            firstCell.SetSharedFormula("=B3*2", 10, 1);

            // Calculate formulas so that the results are stored in the cells
            workbook.CalculateFormula();

            // Optional: display the resulting values in the console for verification
            for (int i = 2; i < 12; i++)
            {
                Console.WriteLine($"C{i + 1} = {cells[i, 2].Value}");
            }

            // Save the workbook
            workbook.Save("SharedFormulaDemo.xlsx");
        }
    }
}
