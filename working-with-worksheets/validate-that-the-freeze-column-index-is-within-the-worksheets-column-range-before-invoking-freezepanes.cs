using System;
using Aspose.Cells;

namespace FreezePaneValidationDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Desired freeze pane parameters
            int freezeRow = 5;          // zero‑based row index where the pane will be split
            int freezeColumn = 300;    // zero‑based column index to validate
            int frozenRows = 5;        // number of rows to keep visible in the top pane
            int frozenColumns = 5;     // number of columns to keep visible in the left pane

            // Validate that the column index is within the worksheet's column range
            // MaxColumn is zero‑based and depends on the workbook format (e.g., 255 for XLS)
            int maxColumn = workbook.Settings.MaxColumn;

            if (freezeColumn < 0 || freezeColumn > maxColumn)
            {
                Console.WriteLine($"Requested freeze column {freezeColumn} is out of range (0‑{maxColumn}).");
                // Adjust to the maximum allowed column index
                freezeColumn = maxColumn;
                Console.WriteLine($"Freeze column adjusted to {freezeColumn}.");
            }

            // Freeze panes using the validated indices (FreezePanes rule)
            worksheet.FreezePanes(freezeRow, freezeColumn, frozenRows, frozenColumns);
            Console.WriteLine($"Panes frozen at row {freezeRow}, column {freezeColumn}.");

            // Save the workbook (save rule)
            workbook.Save("FreezePaneValidated.xlsx");
            Console.WriteLine("Workbook saved as FreezePaneValidated.xlsx");
        }
    }
}