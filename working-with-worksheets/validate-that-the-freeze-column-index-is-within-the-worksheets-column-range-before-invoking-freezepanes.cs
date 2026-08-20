// Title: Validate Freeze Column Index Before Using FreezePanes in Aspose.Cells for .NET
// Description: Demonstrates how to retrieve the workbook's maximum column (Workbook.Settings.MaxColumn), verify that a requested freezeColumn is within that range, adjust it if necessary, ensure frozenColumns does not exceed the split point, and then apply Worksheet.FreezePanes safely before saving the file.
// Keywords: Aspose.Cells FreezePanes validation | C# max column limit | worksheet freeze column range | adjust freezeColumn Aspose | prevent FreezePanes error
// Common Searches: Aspose.Cells check freeze column against MaxColumn | C# validate FreezePanes indices | how to avoid column overflow when freezing panes | adjust frozen columns to split point Aspose.Cells | freeze panes beyond Excel column limit
// Developer Intent: Make sure the column index supplied to Worksheet.FreezePanes does not exceed the worksheet's allowed column range.
// Use Cases: Clamp a user‑provided freezeColumn to Workbook.Settings.MaxColumn to prevent runtime exceptions. | Synchronize frozenColumns with the validated freezeColumn so the left pane remains consistent. | Apply the corrected indices in FreezePanes and generate a compliant Excel file.
// AI Prompts: Create a reusable C# method that checks freezeRow, freezeColumn, frozenRows, and frozenColumns against workbook limits and returns safe values for Worksheet.FreezePanes. | Generate code that logs a warning and automatically reduces freezeColumn when it is greater than Workbook.Settings.MaxColumn in Aspose.Cells. | Show an example handling different Excel version column limits (e.g., 255 vs 16383) while freezing panes with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace FreezePanesValidationDemo
{
    // Demonstrates how to retrieve the workbook's maximum column (Workbook.Settings.MaxColumn), verify that a requested freezeColumn is within that range, adjust it if necessary, ensure frozenColumns does not exceed the split point, and then apply Worksheet.FreezePanes safely before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Desired freeze position (zero‑based indices)
            int freezeRow = 5;      // row index where the pane will be split
            int freezeColumn = 300; // column index to validate
            int frozenRows = 5;     // number of rows to keep visible in the top pane
            int frozenColumns = 5;  // number of columns to keep visible in the left pane

            // Get the maximum column index allowed by the workbook settings
            int maxColumn = workbook.Settings.MaxColumn; // zero‑based, e.g., 255 for Excel97‑2003

            // Validate the column index before calling FreezePanes
            if (freezeColumn > maxColumn)
            {
                Console.WriteLine($"Requested column index {freezeColumn} exceeds the maximum column {maxColumn}. Adjusting to {maxColumn}.");
                freezeColumn = maxColumn;
            }

            // Ensure that the number of frozen columns does not exceed the freeze column index
            if (frozenColumns > freezeColumn)
            {
                Console.WriteLine($"Requested frozen columns {frozenColumns} exceed the freeze column index {freezeColumn}. Adjusting to {freezeColumn}.");
                frozenColumns = freezeColumn;
            }

            // Freeze panes using the validated indices
            worksheet.FreezePanes(freezeRow, freezeColumn, frozenRows, frozenColumns);

            // Save the workbook (lifecycle rule: save)
            workbook.Save("FreezePanesValidated.xlsx");
        }
    }
}
