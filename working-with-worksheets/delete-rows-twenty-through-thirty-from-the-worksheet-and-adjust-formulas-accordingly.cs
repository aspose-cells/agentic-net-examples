using System;
using Aspose.Cells;

class DeleteRowsExample
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (you can change the index or name as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Delete rows 20 through 30 (Excel rows are 1‑based, Cells API is 0‑based)
        // Row index for row 20 is 19, and we need to delete 11 rows (20‑30 inclusive)
        // The third parameter 'true' updates all formula references after deletion
        worksheet.Cells.DeleteRows(19, 11, true);

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}