// Title: Insert a SUM formula into each row of a defined range using Aspose.Cells for .NET (C#)
// Common Searches: how to add a SUM formula to every row in a range with Aspose.Cells C# | Aspose.Cells loop to set column D formula based on columns A to C | C# generate row‑wise SUM formulas programmatically in Excel using Aspose | macro‑like iteration over cells to assign formulas in Aspose.Cells | create dynamic SUM expressions for each row with Aspose.Cells .NET
// Tags: row-wise SUM formula insertion Aspose.Cells | iterate over cell range set formula C# | dynamic Excel formula generation Aspose.Cells | assign column D sum of columns A to C .NET | macro style formula loop Aspose.Cells

using System;
using Aspose.Cells;

// Creates a new workbook, defines a row range, loops through each row, builds cell addresses for columns A‑C, assigns a SUM formula to column D, and saves the file as SumFormulaRows.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook.
        Workbook workbook = new Workbook();

        // Access the first worksheet.
        Worksheet sheet = workbook.Worksheets[0];

        // Define the range over which to insert SUM formulas.
        // Example: rows 1 to 10 (0‑based index 0‑9), columns A to C (0‑2).
        int startRow = 0;      // Row 1
        int endRow = 9;        // Row 10
        int startCol = 0;      // Column A
        int endCol = 2;        // Column C
        int formulaCol = 3;    // Column D where the SUM formula will be placed

        // Iterate over each row in the defined range.
        for (int row = startRow; row <= endRow; row++)
        {
            // Build the address of the first and last cells to sum (e.g., A1:C1).
            string firstCell = CellsHelper.CellIndexToName(startCol, row);
            string lastCell  = CellsHelper.CellIndexToName(endCol, row);

            // Set the SUM formula in the target column.
            // The formula will look like: =SUM(A1:C1)
            sheet.Cells[row, formulaCol].Formula = $"=SUM({firstCell}:{lastCell})";
        }

        // Save the workbook to a file.
        workbook.Save("SumFormulaRows.xlsx");
    }
}
