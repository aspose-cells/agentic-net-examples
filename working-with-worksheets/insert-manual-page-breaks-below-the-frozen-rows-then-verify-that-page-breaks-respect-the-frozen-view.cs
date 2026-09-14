// Title: Insert a manual horizontal page break below frozen rows and verify freeze pane settings using Aspose.Cells for .NET
// AI Prompts: Add a horizontal page break at row 5 after freezing the top rows and output its Row, StartColumn, and EndColumn values. | Call GetFreezedPanes to obtain the freeze pane coordinates and counts, then display them in the console. | Save the workbook and open it to confirm that the page break starts immediately after the frozen area.
// Common Searches: Aspose.Cells how to add a page break after frozen rows in C# | C# example to retrieve freeze pane parameters after inserting manual page breaks | using FreezePanes and HorizontalPageBreaks together with Aspose.Cells .NET | verify that printed pages start after frozen rows using Aspose.Cells | list all horizontal page breaks and frozen pane details in an Aspose.Cells worksheet
// Tags: horizontal page break after FreezePanes Aspose.Cells | retrieve freeze pane parameters C# | validate page break placement with frozen rows | Aspose.Cells worksheet page break verification | freeze panes and manual page breaks .NET

using System;
using Aspose.Cells;

namespace AsposeCellsPageBreakDemo
{
    // The sample creates a workbook, fills 50 rows, freezes the first five rows and the first column, inserts a horizontal page break just below the frozen rows, prints freeze pane and page break details to the console, and saves the file as PageBreaksWithFreeze.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate the worksheet with sample data (e.g., 50 rows)
            for (int i = 0; i < 50; i++)
            {
                worksheet.Cells[i, 0].PutValue($"Row {i + 1}");
            }

            // Freeze the top 5 rows (rows 0‑4) and the first column (column 0)
            // Parameters: row index, column index, number of frozen rows, number of frozen columns
            worksheet.FreezePanes(5, 1, 5, 1);

            // Add a manual horizontal page break just below the frozen rows (at row index 5)
            // This places the break after the frozen area so that printing starts from the next visible row
            worksheet.HorizontalPageBreaks.Add(5);

            // Verify freeze pane information
            int freezeRow, freezeColumn, frozenRows, frozenColumns;
            bool hasFreeze = worksheet.GetFreezedPanes(out freezeRow, out freezeColumn, out frozenRows, out frozenColumns);
            Console.WriteLine($"Worksheet has freeze panes: {hasFreeze}");
            if (hasFreeze)
            {
                Console.WriteLine($"Freeze position - Row: {freezeRow}, Column: {freezeColumn}");
                Console.WriteLine($"Frozen rows: {frozenRows}, Frozen columns: {frozenColumns}");
            }

            // Verify the added horizontal page break
            Console.WriteLine("Horizontal Page Breaks:");
            foreach (HorizontalPageBreak hpb in worksheet.HorizontalPageBreaks)
            {
                Console.WriteLine($"Break at Row: {hpb.Row}, StartColumn: {hpb.StartColumn}, EndColumn: {hpb.EndColumn}");
            }

            // Save the workbook to verify the result in Excel
            workbook.Save("PageBreaksWithFreeze.xlsx");
        }
    }
}
