// Title: C# – Detect Frozen Panes in an Aspose.Cells Worksheet (IsFreezePanes & GetFreezedPanes)
// Description: The sample creates or loads a workbook, applies FreezePanes at cell C5, then shows how to verify frozen panes using the IsFreezePanes property and the GetFreezedPanes method. It prints the freeze position, the number of frozen rows and columns, and saves the workbook.
// Keywords: Aspose.Cells frozen panes C# | IsFreezePanes property | GetFreezedPanes method | detect freeze panes .NET | worksheet FreezePanes example | Aspose.Cells API check freeze | C# Excel freeze pane detection
// Common Searches: Aspose.Cells check if worksheet has frozen panes | C# IsFreezePanes property usage | GetFreezedPanes returns false | How to read freeze pane coordinates Aspose.Cells | Detect frozen rows and columns in .NET Excel library
// Developer Intent: Determine whether a worksheet contains frozen panes and obtain their row/column coordinates.
// Use Cases: Validate a workbook before publishing to ensure no unintended frozen sections. | Adjust UI layout dynamically based on the presence and location of frozen rows or columns. | Log freeze pane settings for auditing, debugging, or migration scripts.
// AI Prompts: Write C# code using Aspose.Cells that checks for frozen panes with IsFreezePanes and prints the freeze row, column, and counts. | Create a reusable method that returns a boolean for frozen panes and outputs the coordinates using GetFreezedPanes. | Show how to combine IsFreezePanes and GetFreezedPanes to both detect and detail frozen panes in an Aspose.Cells worksheet.

using System;
using Aspose.Cells;

namespace AsposeCellsFreezePaneCheck
{
    // The sample creates or loads a workbook, applies FreezePanes at cell C5, then shows how to verify frozen panes using the IsFreezePanes property and the GetFreezedPanes method. It prints the freeze position, the number of frozen rows and columns, and saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // creates an empty workbook

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Example: freeze panes at cell C5 (row index 4, column index 2) with 2 frozen rows and 1 frozen column
            worksheet.FreezePanes(4, 2, 2, 1);

            // Variables to receive freeze pane details
            int freezeRow, freezeColumn, frozenRows, frozenColumns;

            // GetFreezedPanes returns true if the worksheet has frozen panes
            bool hasFrozenPanes = worksheet.GetFreezedPanes(
                out freezeRow,
                out freezeColumn,
                out frozenRows,
                out frozenColumns);

            // Output the result
            Console.WriteLine($"Worksheet has frozen panes: {hasFrozenPanes}");
            if (hasFrozenPanes)
            {
                Console.WriteLine($"Freeze position - Row: {freezeRow}, Column: {freezeColumn}");
                Console.WriteLine($"Frozen rows: {frozenRows}, Frozen columns: {frozenColumns}");
            }

            // Save the workbook (optional, just to demonstrate lifecycle usage)
            workbook.Save("FreezePaneCheckOutput.xlsx");
        }
    }
}
