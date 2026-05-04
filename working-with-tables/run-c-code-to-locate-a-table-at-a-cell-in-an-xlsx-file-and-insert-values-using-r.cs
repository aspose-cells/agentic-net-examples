using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class InsertValuesIntoTable
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index or name as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Define the cell that should be inside the target table (e.g., "B2")
        string anchorCellName = "B2";
        Cell anchorCell = worksheet.Cells[anchorCellName];

        // Locate the ListObject (table) that contains the anchor cell
        ListObject targetTable = null;
        foreach (ListObject table in worksheet.ListObjects)
        {
            int startRow = table.StartRow;
            int startColumn = table.StartColumn;
            int endRow = table.EndRow;
            int endColumn = table.EndColumn;

            if (anchorCell.Row >= startRow && anchorCell.Row <= endRow &&
                anchorCell.Column >= startColumn && anchorCell.Column <= endColumn)
            {
                targetTable = table;
                break;
            }
        }

        if (targetTable == null)
        {
            Console.WriteLine($"No table found containing cell {anchorCellName}.");
            return;
        }

        // Example: Insert values relative to the top‑left data cell of the table
        // Row/column offsets are zero‑based relative to the first data row/column
        // (not counting the header row)
        int rowOffset = 2;      // third data row (0 = first data row)
        int columnOffset = 1;   // second data column

        // Insert a string value
        targetTable.PutCellValue(rowOffset, columnOffset, "Inserted Text");

        // Insert a numeric value in another cell (e.g., first data row, third column)
        targetTable.PutCellValue(0, 2, 12345);

        // Save the modified workbook (replace with desired output path)
        workbook.Save("output.xlsx");
    }
}