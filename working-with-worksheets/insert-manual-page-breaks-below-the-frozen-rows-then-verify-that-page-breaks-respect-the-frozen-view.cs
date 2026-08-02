// Title: Aspose.Cells for .NET – Add Horizontal Page Breaks Below Frozen Rows and Verify Placement (C#)
// Description: This C# example shows how to create a workbook with Aspose.Cells, freeze the top five rows and the first column, insert a manual horizontal page break right after the frozen rows (plus extra breaks every 20 rows), and programmatically validate that every break lies below the frozen area before saving the file.
// Keywords: Aspose.Cells | C# | horizontal page break | freeze panes | manual page break | worksheet pagination | Excel export | page break validation | frozen rows | print layout
// Common Searches: Aspose.Cells add page break after frozen rows | C# insert horizontal page break below freeze panes | verify page break position with Aspose.Cells | how to prevent page breaks inside frozen panes Aspose | Aspose.Cells pagination with frozen header
// Developer Intent: Insert horizontal page breaks after frozen rows and confirm they are positioned outside the frozen pane.
// Use Cases: Create printable reports where header rows stay frozen and each printed page starts after the header. | Generate PDFs from Excel worksheets with consistent page breaks that respect frozen panes. | Automate layout checks before exporting to ensure page breaks never intersect frozen areas.
// AI Prompts: Write C# code using Aspose.Cells to freeze the first 5 rows and column A, then add a horizontal page break immediately after the frozen rows. | Provide a method that iterates through all HorizontalPageBreak objects in a worksheet and logs whether each break is below the frozen rows. | Show how to save the workbook and open it in Excel to visually confirm that page breaks appear after the frozen pane.

using System;
using Aspose.Cells;

namespace AsposeCellsPageBreakDemo
{
    // This C# example shows how to create a workbook with Aspose.Cells, freeze the top five rows and the first column, insert a manual horizontal page break right after the frozen rows (plus extra breaks every 20 rows), and programmatically validate that every break lies below the frozen area before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate the worksheet with sample data (100 rows, 3 columns)
            for (int i = 0; i < 100; i++)
            {
                worksheet.Cells[i, 0].PutValue($"Row {i + 1}");
                worksheet.Cells[i, 1].PutValue(i * 10);
                worksheet.Cells[i, 2].PutValue(DateTime.Today.AddDays(i).ToShortDateString());
            }

            // Freeze the top 5 rows and the first column
            int freezeRow = 5;      // zero‑based index of the row where the freeze starts
            int freezeColumn = 1;   // zero‑based index of the column where the freeze starts
            worksheet.FreezePanes(freezeRow, freezeColumn, freezeRow, freezeColumn);

            // Insert a manual horizontal page break just below the frozen rows
            // Row index is zero‑based; adding at 'freezeRow' places the break after the frozen area
            worksheet.HorizontalPageBreaks.Add(freezeRow);

            // Add additional page breaks every 20 rows after the frozen area for demonstration
            for (int row = freezeRow + 20; row < 100; row += 20)
            {
                worksheet.HorizontalPageBreaks.Add(row);
            }

            // Verify freeze pane information
            bool hasFreeze = worksheet.GetFreezedPanes(out int frozenPosRow, out int frozenPosColumn,
                                                       out int frozenRows, out int frozenColumns);
            Console.WriteLine($"Worksheet has frozen panes: {hasFreeze}");
            if (hasFreeze)
            {
                Console.WriteLine($"Freeze position - Row: {frozenPosRow}, Column: {frozenPosColumn}");
                Console.WriteLine($"Number of frozen rows: {frozenRows}, frozen columns: {frozenColumns}");
            }

            // Verify that each manual page break is positioned below the frozen rows
            Console.WriteLine("\nManual Horizontal Page Breaks (row indices, zero‑based):");
            foreach (HorizontalPageBreak hpb in worksheet.HorizontalPageBreaks)
            {
                Console.WriteLine($"Page break at row {hpb.Row}");
                if (hpb.Row <= frozenRows - 1)
                {
                    Console.WriteLine("  -> Warning: Page break is within the frozen area!");
                }
                else
                {
                    Console.WriteLine("  -> Correct: Page break is below the frozen rows.");
                }
            }

            // Save the workbook to verify the result in Excel
            workbook.Save("PageBreaksWithFrozenRows.xlsx");
        }
    }
}
