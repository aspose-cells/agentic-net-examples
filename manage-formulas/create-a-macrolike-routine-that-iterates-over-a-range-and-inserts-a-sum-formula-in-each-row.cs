// Title: Insert a row‑wise SUM(A:D) formula into column E using Aspose.Cells for .NET (C#)
// Common Searches: how to add a SUM formula to each row with Aspose.Cells C# | Aspose.Cells loop to set formulas in a column for a range of rows | C# generate dynamic SUM(A2:D2) formulas using Aspose.Cells | populate Excel column with row totals using Aspose.Cells .NET | write macro‑like code to insert SUM formulas across rows in Aspose.Cells
// Tags: Aspose.Cells set cell formula programmatically | C# insert SUM formula into Excel column | iterate rows to assign formulas Aspose.Cells | dynamic Excel formula generation .NET | populate column with aggregate formulas using Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a workbook, fills columns A‑D with sample data for rows 2‑10, then iterates each row to place a SUM(Ax:Dx) formula in column E, and saves the file as SumFormulaExample.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Fill sample data in columns A to D for rows 2 to 10
        for (int row = 1; row <= 9; row++) // zero‑based index (row 1 = Excel row 2)
        {
            for (int col = 0; col < 4; col++) // columns A‑D
            {
                sheet.Cells[row, col].PutValue(row + col);
            }
        }

        // Define the range where SUM formulas will be inserted (rows 2‑10, column E)
        int startRow = 1; // Excel row 2
        int endRow = 9;   // Excel row 10
        int sumColumn = 4; // Column E (zero‑based)

        // Iterate over each row and set the SUM formula for columns A‑D
        for (int r = startRow; r <= endRow; r++)
        {
            // Build the formula string, e.g., =SUM(A2:D2)
            string formula = string.Format("=SUM(A{0}:D{0})", r + 1);
            sheet.Cells[r, sumColumn].Formula = formula;
        }

        // Save the workbook
        workbook.Save("SumFormulaExample.xlsx");
    }
}
