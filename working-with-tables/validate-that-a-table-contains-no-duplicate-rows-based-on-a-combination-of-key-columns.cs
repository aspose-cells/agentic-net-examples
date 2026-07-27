using System;
using Aspose.Cells;

class ValidateDuplicatesDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // ----- Sample data (including header row) -----
        // Columns: ID (0), Name (1), Age (2)
        cells["A1"].PutValue("ID");
        cells["B1"].PutValue("Name");
        cells["C1"].PutValue("Age");

        cells["A2"].PutValue(1);
        cells["B2"].PutValue("John");
        cells["C2"].PutValue(30);

        cells["A3"].PutValue(2);
        cells["B3"].PutValue("Jane");
        cells["C3"].PutValue(25);

        // Duplicate row based on ID and Name
        cells["A4"].PutValue(1);
        cells["B4"].PutValue("John");
        cells["C4"].PutValue(35);

        cells["A5"].PutValue(3);
        cells["B5"].PutValue("Bob");
        cells["C5"].PutValue(40);
        // ----------------------------------------------

        // Define the range that contains the table (including header)
        int startRow = 0;                     // first row (header)
        int startColumn = 0;                  // first column (A)
        int endRow = cells.MaxDataRow;        // last row with data
        int endColumn = cells.MaxDataColumn;  // last column with data

        // Count rows before duplicate removal (header + data rows)
        int rowsBefore = endRow - startRow + 1;

        // Remove duplicates based on key columns: ID (offset 0) and Name (offset 1)
        // hasHeaders = true because the first row contains column names
        cells.RemoveDuplicates(startRow, startColumn, endRow, endColumn, true, new int[] { 0, 1 });

        // Count rows after duplicate removal
        int rowsAfter = cells.MaxDataRow - startRow + 1;

        // Validation result
        if (rowsBefore == rowsAfter)
        {
            Console.WriteLine("No duplicate rows found based on the key columns.");
        }
        else
        {
            Console.WriteLine($"Duplicates detected. Rows before: {rowsBefore}, after removing duplicates: {rowsAfter}");
        }

        // Save the workbook (optional, shows the table after duplicate removal)
        workbook.Save("ValidateDuplicatesDemo.xlsx");
    }
}