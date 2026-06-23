using System;
using Aspose.Cells;

namespace AsposeCellsFreezePaneExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data in columns A to E (indices 0‑4)
            for (int col = 0; col < 5; col++)
            {
                cells[0, col].PutValue($"Header {(char)('A' + col)}");
                cells[1, col].PutValue($"Data {(char)('A' + col)}1");
                cells[2, col].PutValue($"Data {(char)('A' + col)}2");
            }

            // Desired freeze area: row 2, column 4 (E column, zero‑based index 4)
            // We will delete the first two columns (A and B) before freezing.
            // After deletion, the original column index 4 becomes 2.
            int columnsToDelete = 2; // number of columns to remove from the left

            // Delete the first two columns (indices 0 and 1)
            // Using DeleteColumns to remove a range and update references
            cells.DeleteColumns(0, columnsToDelete, true);

            // Freeze panes using the updated column index (original 4 - 2 = 2)
            int freezeRow = 2;      // freeze after row 2 (zero‑based)
            int freezeColumn = 4 - columnsToDelete; // adjusted column index
            int frozenRows = freezeRow;   // number of rows to keep visible on top
            int frozenColumns = freezeColumn; // number of columns to keep visible on left

            sheet.FreezePanes(freezeRow, freezeColumn, frozenRows, frozenColumns);

            // Save the workbook
            workbook.Save("FreezePaneAfterDeleteColumns.xlsx");
        }
    }
}