// Title: C# – Duplicate a Worksheet, Assign a New TabId, and Clear Data While Preserving Formatting with Aspose.Cells
// Description: Load a workbook, copy a specific worksheet using the AddCopy method, give the copy a unique TabId, and remove all cell values without affecting styles by clearing the used range with ClearContents. Finally, save the updated file. This example shows how to keep layout and formatting intact while resetting data.
// Keywords: Aspose.Cells copy worksheet C# | AddCopy method | set worksheet TabId | clear cell contents preserve formatting | ClearContents used range | duplicate sheet Aspose.Cells | C# Excel workbook manipulation
// Common Searches: Aspose.Cells duplicate worksheet C# | How to set TabId for copied sheet in Aspose.Cells | Clear data but keep formatting Aspose.Cells | AddCopy example Aspose.Cells .NET | Reset worksheet values while preserving styles
// Developer Intent: Create a copy of an existing worksheet, give it a distinct TabId, and erase its data while leaving all formatting unchanged.
// Use Cases: Generate a clean template from a master sheet for recurring reports. | Prepare a user‑input sheet that shares the layout of a reference sheet but has a separate TabId for navigation. | Refresh data in a duplicated worksheet without rebuilding the style definitions.
// AI Prompts: Write C# code with Aspose.Cells that copies a worksheet, assigns a new TabId, and clears only the cell values while keeping all formatting. | Show an Aspose.Cells .NET example that uses AddCopy, sets a non‑conflicting TabId, and calls ClearContents on the used range. | Explain the steps to duplicate a sheet, change its TabId, and remove data without affecting styles in Aspose.Cells.

using System;
using Aspose.Cells;

namespace WorksheetDuplicationDemo
{
    // Load a workbook, copy a specific worksheet using the AddCopy method, give the copy a unique TabId, and remove all cell values without affecting styles by clearing the used range with ClearContents. Finally, save the updated file. This example shows how to keep layout and formatting intact while resetting data.
    class Program
    {
        static void Main()
        {
            // Load the source workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Name of the worksheet to duplicate
            string sourceSheetName = "Sheet1";

            // Add a copy of the source worksheet to the workbook
            // The AddCopy method returns the index of the newly created sheet
            int copiedIndex = workbook.Worksheets.AddCopy(sourceSheetName);
            Worksheet copiedSheet = workbook.Worksheets[copiedIndex];

            // Assign a new TabId to the duplicated worksheet
            // Here we simply set it to a value that does not conflict with existing ones
            // (e.g., original TabId + 1000). Adjust as needed.
            copiedSheet.TabId = workbook.Worksheets[sourceSheetName].TabId + 1000;

            // Clear all cell values while preserving formatting
            // ClearContents removes only the data, leaving styles intact.
            // Use the used range to avoid processing the entire possible sheet size.
            int maxRow = copiedSheet.Cells.MaxDataRow;
            int maxColumn = copiedSheet.Cells.MaxDataColumn;
            if (maxRow >= 0 && maxColumn >= 0) // Ensure there is at least one used cell
            {
                copiedSheet.Cells.ClearContents(0, 0, maxRow, maxColumn);
            }

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx");
        }
    }
}
