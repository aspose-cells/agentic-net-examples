using Aspose.Cells;

class CopyRowsPreserveHidden
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data in rows 0‑4
        for (int i = 0; i < 5; i++)
        {
            cells[i, 0].PutValue($"Row {i + 1}");
        }

        // Hide specific rows (row indices are zero‑based)
        cells.HideRow(1); // hide row 2
        cells.HideRow(3); // hide row 4

        // Define the range of rows to copy
        int sourceRowIndex = 0;          // start copying from row 0
        int rowCount = 5;                // number of rows to copy
        int destinationRowIndex = cells.MaxRow + 1; // insert after existing rows

        // Insert empty rows at the destination to make space
        cells.InsertRows(destinationRowIndex, rowCount);

        // Copy rows data and formats
        cells.CopyRows(cells, sourceRowIndex, destinationRowIndex, rowCount);

        // Preserve hidden state and other row settings
        for (int i = 0; i < rowCount; i++)
        {
            Row sourceRow = cells.Rows[sourceRowIndex + i];
            Row destRow = cells.Rows[destinationRowIndex + i];
            // Copy settings (including IsHidden) from source to destination row
            destRow.CopySettings(sourceRow, false);
        }

        // Save the workbook
        workbook.Save("CopyRowsPreserveHidden.xlsx");
    }
}