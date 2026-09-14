// Title: Add row‑specific multiplication formulas to column B while iterating column A using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that reads each value in column A and writes a formula `=A{row}*{row}` into the same row of column B. | Show how to invoke workbook.CalculateFormula() and persist the file after adding the row‑dependent formulas using Aspose.Cells.
// Common Searches: Aspose.Cells C# loop through column A and set formula in column B based on row number | How to create a multiplication formula that uses the current row index in Aspose.Cells | Programmatically add dynamic formulas to an Excel worksheet with Aspose.Cells .NET | How to evaluate all formulas and write the Excel file when using Aspose.Cells in C# | Generate row‑wise formulas in Excel using the Aspose.Cells library
// Tags: Aspose.Cells insert row‑based formula | C# generate Excel multiplication formula | Aspose.Cells calculate workbook formulas | C# save Excel file with Aspose.Cells | loop cells column A Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new workbook, fills A1:A10 with values 1‑10, iterates each row to place a formula in column B that multiplies the A‑cell value by its row number, calculates all formulas, and saves the result as MultiplicationResults.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate column A with sample values (1 to 10)
            for (int row = 0; row < 10; row++)
            {
                // Row index is zero‑based; cell "A1" corresponds to row 0, column 0
                cells[row, 0].PutValue(row + 1);
            }

            // Loop through column A and assign incremental row references in a multiplication formula
            // The result will be stored in column B of the same row
            for (int row = 0; row < 10; row++)
            {
                // Build the formula string: =A{rowNumber}*{rowNumber}
                // Excel rows are 1‑based, so add 1 to the zero‑based index
                int excelRow = row + 1;
                string formula = $"=A{excelRow}*{excelRow}";

                // Place the formula in column B (index 1) of the current row
                cells[row, 1].Formula = formula;
            }

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Save the workbook to a file (lifecycle rule: save)
            workbook.Save("MultiplicationResults.xlsx");
        }
    }
}
