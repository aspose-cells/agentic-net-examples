// Title: C# Example: Insert Horizontal Page Breaks Below Frozen Rows with Aspose.Cells for .NET
// Description: Demonstrates how to freeze the top rows and first column of a worksheet, add horizontal page breaks immediately after the frozen area (and at additional rows), list the freeze‑pane settings and page‑break indices, and save the workbook as PageBreaksWithFrozenRows.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# | horizontal page break | freeze panes | worksheet pagination | Excel page break API | .NET Excel export | verify page break position | GitHub Aspose.Cells example | programmatic page break
// Common Searches: add page break after frozen rows Aspose.Cells .NET | C# freeze panes and insert horizontal page break | how to list page break rows with Aspose.Cells | verify page break placement relative to freeze panes | Aspose.Cells example for pagination with frozen headers
// Developer Intent: Insert horizontal page breaks right after frozen rows and confirm their locations programmatically.
// Use Cases: Create printable Excel reports where header rows stay frozen while each printed page starts after the header. | Generate large workbooks with consistent pagination and frozen navigation panes for easier data review. | Debug worksheet layout by outputting freeze‑pane details and the exact row indices of all horizontal page breaks.
// AI Prompts: Generate C# code using Aspose.Cells to freeze the first N rows and column, then add a horizontal page break directly below the frozen area and list all page‑break rows. | Show how to retrieve freeze‑pane information and verify that added horizontal page breaks are positioned correctly in an Aspose.Cells worksheet. | Explain best practices for preventing page breaks from intersecting frozen rows when adding them programmatically with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsPageBreakDemo
{
    // Demonstrates how to freeze the top rows and first column of a worksheet, add horizontal page breaks immediately after the frozen area (and at additional rows), list the freeze‑pane settings and page‑break indices, and save the workbook as PageBreaksWithFrozenRows.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data (30 rows, 5 columns)
            for (int i = 0; i < 30; i++)
            {
                for (int j = 0; j < 5; j++)
                {
                    worksheet.Cells[i, j].PutValue($"R{i + 1}C{j + 1}");
                }
            }

            // Freeze the top 5 rows (rows 0-4) and the first column (column 0)
            int freezeRow = 5;          // Row index where the freeze line starts (zero‑based)
            int freezeColumn = 1;       // Column index where the freeze line starts (zero‑based)
            int frozenRows = 5;         // Number of rows to freeze
            int frozenColumns = 1;      // Number of columns to freeze
            worksheet.FreezePanes(freezeRow, freezeColumn, frozenRows, frozenColumns);

            // Add a horizontal page break immediately below the frozen rows
            // Since rows are zero‑based, the first row after the frozen area is index 5
            worksheet.HorizontalPageBreaks.Add(frozenRows);

            // Add additional page breaks for demonstration
            worksheet.HorizontalPageBreaks.Add(15); // After row 15
            worksheet.HorizontalPageBreaks.Add(25); // After row 25

            // Verify freeze pane information
            bool hasFreeze = worksheet.GetFreezedPanes(out int fpRow, out int fpColumn, out int fpRows, out int fpColumns);
            Console.WriteLine($"Worksheet has freeze panes: {hasFreeze}");
            if (hasFreeze)
            {
                Console.WriteLine($"Freeze position - Row: {fpRow}, Column: {fpColumn}");
                Console.WriteLine($"Frozen rows: {fpRows}, Frozen columns: {fpColumns}");
            }

            // Verify that page breaks are positioned correctly relative to frozen rows
            Console.WriteLine("Horizontal page breaks (row indices):");
            foreach (HorizontalPageBreak hpb in worksheet.HorizontalPageBreaks)
            {
                Console.WriteLine($"Row: {hpb.Row}");
            }

            // Save the workbook
            workbook.Save("PageBreaksWithFrozenRows.xlsx");
        }
    }
}
