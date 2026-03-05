using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Get a cell that belongs to the table (e.g., A2)
        Cell cell = worksheet.Cells["A2"];

        // Retrieve the table (ListObject) that contains this cell
        ListObject table = cell.GetTable();

        if (table != null)
        {
            // Add values to the third data row of the table (rowOffset = 2)
            // Column offsets are zero‑based relative to the table's first column
            table.PutCellValue(2, 0, 123);                     // First column
            table.PutCellValue(2, 1, "New Item");              // Second column
            table.PutCellValue(2, 2, DateTime.Now);            // Third column
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}