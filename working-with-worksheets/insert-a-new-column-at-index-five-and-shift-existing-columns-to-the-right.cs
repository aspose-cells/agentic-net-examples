// Title: C# – Insert a Column at Index 5 with Aspose.Cells (Shift Existing Columns)
// Description: Demonstrates how to add a column at zero‑based index 5 in a worksheet using Aspose.Cells for .NET, automatically shifting all subsequent columns to the right, adding a header, and saving the file as XLSX.
// Keywords: Aspose.Cells InsertColumn C# | add column index 5 Aspose | shift columns right .NET | Aspose.Cells worksheet column insertion | C# Excel column manipulation
// Common Searches: Aspose.Cells insert column at position 5 | C# insert column and shift existing columns in Excel | how to add a new column in Aspose.Cells worksheet | InsertColumn method zero‑based index example
// Developer Intent: Add a column at the sixth position (index 5) and move all later columns one place to the right.
// Use Cases: Introduce a new data field without overwriting current columns. | Create a placeholder for future calculations while preserving layout. | Insert a header column before exporting the workbook to Excel.
// AI Prompts: Generate C# code using Aspose.Cells to insert a column at index 5, shift existing columns, and set a header value. | Show how to insert multiple consecutive columns starting at a specific index while keeping existing data intact. | Explain the zero‑based indexing of InsertColumn and how to adjust formulas after the insertion.

using System;
using Aspose.Cells;

// Demonstrates how to add a column at zero‑based index 5 in a worksheet using Aspose.Cells for .NET, automatically shifting all subsequent columns to the right, adding a header, and saving the file as XLSX.
class InsertColumnDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Fill some sample data (optional, just to see the shift effect)
        for (int row = 0; row < 5; row++)
        {
            for (int col = 0; col < 8; col++)
            {
                cells[row, col].PutValue($"R{row}C{col}");
            }
        }

        // Insert a new column at index 5 (zero‑based, i.e., the 6th column)
        // This shifts all existing columns from index 5 onward to the right.
        cells.InsertColumn(5);

        // Add a header to the newly inserted column (optional)
        cells[0, 5].PutValue("New Column");

        // Save the workbook
        workbook.Save("InsertColumnAtIndex5.xlsx");
    }
}
