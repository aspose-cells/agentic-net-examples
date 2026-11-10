// Title: Insert a running total column with a SUM formula for each row using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add running total column to existing worksheet | how to programmatically insert a SUM formula for each row with Aspose.Cells .NET | create macro‑like routine to calculate row totals in Excel using Aspose.Cells | insert new column and set row‑wise SUM formula in C# Aspose.Cells | generate running total column automatically in .NET Excel file
// Tags: insert column with SUM formula Aspose.Cells | row-wise running total calculation Aspose.Cells | dynamic column insertion Excel .NET | auto-generate running total header Aspose.Cells | programmatic row totals using Aspose.Cells

using System;
using Aspose.Cells;

// The example loads an existing workbook, determines the last used column, inserts a new column for running totals, adds a header, then iterates over each data row building a SUM formula that spans from the first data cell to the original last column, assigns the formula to the new column, and saves the updated workbook.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet sheet = workbook.Worksheets[0];

        // Find the last used column (zero‑based index)
        int originalLastCol = sheet.Cells.MaxColumn;

        // Insert a new column after the existing data columns for the running total
        int totalColIndex = originalLastCol + 1;
        sheet.Cells.InsertColumn(totalColIndex);

        // Add a header for the running total column
        sheet.Cells[0, totalColIndex].PutValue("Running Total");

        // Find the last used row (zero‑based index)
        int lastRow = sheet.Cells.MaxDataRow;

        // Assume the first row (row 0) contains headers, data starts at row 1
        for (int row = 1; row <= lastRow; row++)
        {
            // Build a SUM formula that adds all cells in the current row
            // from the first data column (0) to the original last data column
            string startCell = CellsHelper.CellIndexToName(row, 0);               // e.g., A2
            string endCell   = CellsHelper.CellIndexToName(row, originalLastCol); // e.g., D2
            string formula   = $"=SUM({startCell}:{endCell})";

            // Set the formula in the running total column for the current row
            sheet.Cells[row, totalColIndex].Formula = formula;
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
