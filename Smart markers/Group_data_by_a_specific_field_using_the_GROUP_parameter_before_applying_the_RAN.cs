using System;
using Aspose.Cells;

class GroupDataByField
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Define the data range that contains the source data (e.g., A1:C6)
        CellArea dataRange = new CellArea
        {
            StartRow = 0,      // Row 1 (zero‑based)
            StartColumn = 0,   // Column A
            EndRow = 5,        // Row 6
            EndColumn = 2      // Column C
        };

        // Apply subtotal: first specify the GROUP field (column index 0),
        // then the RANGE (dataRange). The subtotal will sum values in column C (index 2).
        cells.Subtotal(
            dataRange,               // RANGE parameter
            0,                       // GROUP parameter – group by first column
            ConsolidationFunction.Sum,
            new int[] { 2 }          // Columns to subtotal
        );

        // Save the workbook with the applied grouping
        workbook.Save("output.xlsx");
    }
}