using System;
using Aspose.Cells;

namespace AsposeCellsRemoveDuplicatesExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Add header row (hasHeaders = true)
            cells["A1"].PutValue("ID");          // Column 0
            cells["B1"].PutValue("Name");        // Column 1
            cells["C1"].PutValue("Score");       // Column 2 (extra column not part of key)

            // Add sample data with duplicate rows based on the composite key (ID, Name)
            // Row 2
            cells["A2"].PutValue(1);
            cells["B2"].PutValue("Alice");
            cells["C2"].PutValue(85);
            // Row 3 (duplicate of row 2 on ID and Name)
            cells["A3"].PutValue(1);
            cells["B3"].PutValue("Alice");
            cells["C3"].PutValue(90);
            // Row 4
            cells["A4"].PutValue(2);
            cells["B4"].PutValue("Bob");
            cells["C4"].PutValue(78);
            // Row 5 (duplicate of row 4 on ID and Name)
            cells["A5"].PutValue(2);
            cells["B5"].PutValue("Bob");
            cells["C5"].PutValue(82);
            // Row 6 (unique)
            cells["A6"].PutValue(3);
            cells["B6"].PutValue("Charlie");
            cells["C6"].PutValue(91);

            // Define the range that contains the data (including header)
            int startRow = 0;          // Header row index
            int startColumn = 0;       // First column (ID)
            int endRow = 5;            // Zero‑based index of last data row (row 6 in Excel)
            int endColumn = 2;         // Last column (Score)

            // Remove duplicates based on the composite key of columns 0 (ID) and 1 (Name)
            // hasHeaders = true indicates that the first row contains column names
            // columnOffsets specifies which columns form the key (0‑based offsets from startColumn)
            cells.RemoveDuplicates(startRow, startColumn, endRow, endColumn, true, new int[] { 0, 1 });

            // Save the workbook to verify the result
            workbook.Save("RemoveDuplicatesCompositeKey.xlsx", SaveFormat.Xlsx);
        }
    }
}