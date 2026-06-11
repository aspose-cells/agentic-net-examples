using System;
using Aspose.Cells;

namespace AsposeCellsPageBreakDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data (30 rows, 3 columns)
            for (int i = 0; i < 30; i++)
            {
                worksheet.Cells[i, 0].PutValue($"Row {i + 1} - Col A");
                worksheet.Cells[i, 1].PutValue($"Row {i + 1} - Col B");
                worksheet.Cells[i, 2].PutValue($"Row {i + 1} - Col C");
            }

            // Freeze the top 5 rows (zero‑based index 5) and first column
            int freezeRow = 5;          // row index where freeze starts
            int freezeColumn = 0;       // column index where freeze starts
            int frozenRows = 5;         // number of rows to freeze
            int frozenColumns = 1;      // number of columns to freeze
            worksheet.FreezePanes(freezeRow, freezeColumn, frozenRows, frozenColumns);

            // Add a manual horizontal page break just below the frozen rows
            // Row index 5 corresponds to the first row after the frozen area
            int pageBreakRow = frozenRows; // zero‑based row index for the break
            worksheet.HorizontalPageBreaks.Add(pageBreakRow);

            // Verify that the page break respects the frozen view
            // Retrieve freeze pane information
            bool hasFreeze = worksheet.GetFreezedPanes(out int fpRow, out int fpColumn,
                                                       out int fpRows, out int fpColumns);

            // Retrieve the added page break object
            HorizontalPageBreak hBreak = worksheet.HorizontalPageBreaks[0];

            // Output verification details
            Console.WriteLine($"Worksheet has frozen panes: {hasFreeze}");
            if (hasFreeze)
            {
                Console.WriteLine($"Frozen rows: {fpRows}, Frozen columns: {fpColumns}");
                Console.WriteLine($"Freeze position - Row: {fpRow}, Column: {fpColumn}");
            }

            Console.WriteLine($"Added horizontal page break at row index: {hBreak.Row}");
            Console.WriteLine($"Page break starts at column: {hBreak.StartColumn}, ends at column: {hBreak.EndColumn}");

            // Check that the page break row is not within the frozen rows
            if (hBreak.Row >= fpRows)
                Console.WriteLine("Page break is correctly placed below the frozen rows.");
            else
                Console.WriteLine("Page break is incorrectly placed within the frozen area.");

            // Save the workbook to verify visually if needed
            workbook.Save("PageBreakWithFreezeDemo.xlsx");
        }
    }
}