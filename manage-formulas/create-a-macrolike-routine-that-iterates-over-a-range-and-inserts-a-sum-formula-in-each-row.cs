// Title: Insert a row‑by‑row SUM formula into a column using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that loops through rows 1‑10 and writes a `=SUM(A1:C1)` style formula into column D for each row. | Create a reusable method in Aspose.Cells that accepts startRow, endRow, startCol, endCol, and targetCol parameters and inserts a SUM formula for each row in the specified range. | Show how to build a macro‑like routine in Aspose.Cells that programmatically adds row‑level total formulas across a worksheet and saves the workbook.
// Common Searches: aspnet how to add a SUM formula to each row with Aspose.Cells C# loop | C# Aspose.Cells insert row total formula in column D for rows 1 to 10 | programmatically generate per‑row SUM formulas using Aspose.Cells .NET | macro style code to set SUM(A:C) for each row in an Excel file with Aspose.Cells
// Tags: Aspose.Cells row-wise SUM formula insertion | C# loop insert Excel formula Aspose.Cells | dynamic range SUM formula generation .NET | macro‑like formula automation Aspose.Cells | Excel workbook row total calculation C#

using System;
using Aspose.Cells;

// The example creates a new workbook, iterates rows 1‑10, and places a `=SUM(A1:C1)`‑style formula in column D for each row, then saves the file as SummedRows.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook.
        Workbook workbook = new Workbook();

        // Access the first worksheet.
        Worksheet sheet = workbook.Worksheets[0];

        // Define the range to process (e.g., rows 1‑10, columns A‑C).
        int startRow = 0;      // zero‑based index (row 1)
        int endRow = 9;        // row 10
        int startCol = 0;      // column A
        int endCol = 2;        // column C
        int formulaCol = 3;    // column D where the SUM formula will be placed

        // Iterate over each row in the range and insert a SUM formula.
        for (int row = startRow; row <= endRow; row++)
        {
            // Build the address of the first and last cells to sum (e.g., A1:C1).
            string firstCell = CellsHelper.CellIndexToName(row, startCol);
            string lastCell  = CellsHelper.CellIndexToName(row, endCol);

            // Set the SUM formula in the target column.
            sheet.Cells[row, formulaCol].Formula = $"=SUM({firstCell}:{lastCell})";
        }

        // Save the workbook to a file.
        workbook.Save("SummedRows.xlsx");
    }
}
