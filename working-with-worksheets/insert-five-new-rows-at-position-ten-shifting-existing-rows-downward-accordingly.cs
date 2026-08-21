// Title: Insert 5 rows at row 10 using Aspose.Cells for .NET (C#) and shift existing rows down
// Description: C# example that creates a workbook, fills rows 1‑15 with sample data, then calls Cells.InsertRows(9, 5) to add five rows at the tenth position (zero‑based index 9). The operation pushes rows 10‑15 downward, after which the new rows are populated and the file is saved as InsertRowsAtPositionTen.xlsx.
// Keywords: Aspose.Cells insert rows C# | Cells.InsertRows example | add rows at specific index .NET | shift worksheet rows down | Aspose.Cells row insertion | C# Excel automation | programmatic row insertion
// Common Searches: Aspose.Cells how to insert rows at a specific index | Insert multiple rows in Excel using C# Aspose.Cells | InsertRows method zero based index example | Add rows before row 10 in Aspose.Cells workbook | Shift rows down when inserting rows Aspose.Cells
// Developer Intent: Add five new rows at worksheet row 10, move existing rows downward, and optionally fill the inserted rows with data.
// Use Cases: Insert blank rows before a header to accommodate new entries without overwriting existing data. | Expand a table by inserting rows before a totals row while preserving formulas and references. | Generate dynamic reports where the number of detail rows varies and must be inserted at a fixed position.
// AI Prompts: Write C# code that inserts N rows at a given zero‑based index with Aspose.Cells and copies the formatting from the preceding row. | Explain the effect of Cells.InsertRows on merged cells, formulas, and named ranges in an Aspose.Cells worksheet. | Provide a step‑by‑step guide to insert rows at row 10, shift existing rows, and then populate the new rows with custom values using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace InsertRowsExample
{
    // C# example that creates a workbook, fills rows 1‑15 with sample data, then calls Cells.InsertRows(9, 5) to add five rows at the tenth position (zero‑based index 9). The operation pushes rows 10‑15 downward, after which the new rows are populated and the file is saved as InsertRowsAtPositionTen.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some sample data in rows 1-15 (zero‑based indices 0‑14)
            for (int row = 0; row < 15; row++)
            {
                cells[row, 0].PutValue($"Row {row + 1} Column A");
                cells[row, 1].PutValue($"Row {row + 1} Column B");
            }

            // Insert five new rows at position ten (row index 9, zero‑based)
            // This shifts existing rows 10‑15 down by five rows
            cells.InsertRows(9, 5);

            // Optionally add data to the newly inserted rows
            for (int i = 0; i < 5; i++)
            {
                int insertedRow = 9 + i;
                cells[insertedRow, 0].PutValue($"Inserted Row {i + 1} Column A");
                cells[insertedRow, 1].PutValue($"Inserted Row {i + 1} Column B");
            }

            // Save the workbook
            workbook.Save("InsertRowsAtPositionTen.xlsx");
        }
    }
}
