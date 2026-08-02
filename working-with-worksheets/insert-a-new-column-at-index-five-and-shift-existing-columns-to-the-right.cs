// Title: C# AspNet.Cells – Add Column at Position 5 and Shift Right
// Description: Shows how to create a workbook, populate columns A‑G with sample data, and call Cells.InsertColumn(5) to insert a column at the zero‑based position 5. The existing columns from index 5 onward move one place to the right, an optional header is added, and the file is saved as InsertColumnAtIndex5.xlsx.
// Keywords: Aspose.Cells | C# | InsertColumn | add column | shift columns right | worksheet manipulation | sample code | GitHub | zero‑based index | Excel automation
// Common Searches: Aspose.Cells insert column C# example | how to add a column at position 5 using Aspose.Cells | shift columns right after inserting column Aspose.Cells .NET | Cells.InsertColumn method usage | insert column in worksheet programmatically
// Developer Intent: Add a column at position 5 in a worksheet while moving all later columns to the right.
// Use Cases: Insert a placeholder between data sets before importing new metrics | Add a header column for an additional data series without reordering existing columns | Automatically expand report templates with extra columns during batch processing
// AI Prompts: Write C# code with Aspose.Cells that inserts a column at index 5, copies the style from column 4, and updates any formulas referencing shifted cells. | Show how to insert a column at position 5 while preserving merged cells, data validation, and conditional formatting in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsInsertColumnDemo
{
    // Shows how to create a workbook, populate columns A‑G with sample data, and call Cells.InsertColumn(5) to insert a column at the zero‑based position 5. The existing columns from index 5 onward move one place to the right, an optional header is added, and the file is saved as InsertColumnAtIndex5.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Fill some sample data in columns A‑G (indices 0‑6)
            for (int row = 0; row < 5; row++)
            {
                cells[row, 0].PutValue($"A{row}");
                cells[row, 1].PutValue($"B{row}");
                cells[row, 2].PutValue($"C{row}");
                cells[row, 3].PutValue($"D{row}");
                cells[row, 4].PutValue($"E{row}");
                cells[row, 5].PutValue($"F{row}");
                cells[row, 6].PutValue($"G{row}");
            }

            // Insert a new column at index 5 (0‑based). Columns 5 and beyond shift right.
            cells.InsertColumn(5);

            // Add a header to the newly inserted column (optional)
            cells[0, 5].PutValue("New Column");

            // Save the workbook
            workbook.Save("InsertColumnAtIndex5.xlsx");
        }
    }
}
