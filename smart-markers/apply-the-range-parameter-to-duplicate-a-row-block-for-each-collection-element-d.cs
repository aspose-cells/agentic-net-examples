using System;
using System.Collections.Generic;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the template workbook (contains a row block that will be duplicated)
        Workbook workbook = new Workbook("Template.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Define the template block that should be copied.
        // Example: rows 2‑4 (zero‑based index 1‑3) form the block.
        int templateStartRow = 1;          // first row of the block to copy
        int blockRowCount = 3;             // how many rows the block contains

        // Sample collection – each element will cause one copy of the block.
        List<object[]> records = new List<object[]>
        {
            new object[] { "John", "Doe", 30 },
            new object[] { "Jane", "Smith", 25 },
            new object[] { "Bob", "Brown", 40 }
        };

        // Insertion point starts just after the original template block.
        int insertRow = templateStartRow + blockRowCount;

        foreach (var record in records)
        {
            // Duplicate the template block at the current insertion point.
            cells.CopyRows(cells, templateStartRow, insertRow, blockRowCount);

            // Populate the first row of the newly copied block with actual data.
            for (int col = 0; col < record.Length; col++)
            {
                cells[insertRow, col].PutValue(record[col]);
            }

            // Advance the insertion point for the next iteration.
            insertRow += blockRowCount;
        }

        // If the original template block should not appear in the final file,
        // uncomment the following line to delete it.
        // worksheet.Cells.DeleteRows(templateStartRow, blockRowCount);

        // Save the modified workbook.
        workbook.Save("Result.xlsx");
    }
}