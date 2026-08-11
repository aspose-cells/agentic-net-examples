// Title: Copy a worksheet, set a unique TabId, and clear cell values while preserving formatting – Aspose.Cells for .NET
// Description: Learn how to duplicate a worksheet in a .xlsx file using Aspose.Cells for .NET, assign a distinct TabId to the copy, and erase all cell contents without affecting styles. The example loads a workbook, uses AddCopy, updates the TabId, clears the used range with ClearContents, and saves the modified file.
// Keywords: Aspose.Cells copy worksheet C# | set worksheet TabId Aspose.Cells | clear cell values keep formatting | AddCopy method .NET | ClearContents used range | C# Excel workbook manipulation | global Aspose.Cells tutorial
// Common Searches: Aspose.Cells duplicate sheet and change TabId | How to clear only data in copied worksheet Aspose.Cells | C# copy worksheet keep formatting Aspose.Cells | Assign new TabId to Excel sheet with Aspose.Cells
// Developer Intent: Programmatically create a worksheet copy, give it a unique TabId, and remove all data while leaving the original formatting untouched.
// Use Cases: Generate a fresh report template by cloning a styled sheet, assigning a new TabId, and clearing previous entries. | Provide end‑users with a clean worksheet that mirrors a master layout, ensuring consistent formatting across new files. | Prepare a worksheet for API‑driven data import by copying a template sheet, setting a unique TabId to avoid conflicts, and wiping existing values.
// AI Prompts: Write C# code with Aspose.Cells to copy a worksheet, assign a unique TabId, and clear all cell values while preserving formatting. | Explain why ClearContents removes data but leaves cell styles unchanged in Aspose.Cells. | Suggest alternative ways to generate a unique TabId for a duplicated worksheet using Aspose.Cells.

using System;
using Aspose.Cells;

// Learn how to duplicate a worksheet in a .xlsx file using Aspose.Cells for .NET, assign a distinct TabId to the copy, and erase all cell contents without affecting styles. The example loads a workbook, uses AddCopy, updates the TabId, clears the used range with ClearContents, and saves the modified file.
class DuplicateWorksheetDemo
{
    static void Main()
    {
        // Load the source workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Index of the worksheet to duplicate (e.g., the first sheet)
        int sourceIndex = 0;
        Worksheet sourceSheet = workbook.Worksheets[sourceIndex];

        // Duplicate the worksheet using AddCopy (creates a copy with the same content and formatting)
        int copiedIndex = workbook.Worksheets.AddCopy(sourceIndex);
        Worksheet copiedSheet = workbook.Worksheets[copiedIndex];

        // Assign a new TabId to the duplicated sheet (ensure it differs from the original)
        copiedSheet.TabId = copiedSheet.Index + 1000; // any unique integer value

        // Clear all cell values while preserving formatting
        Cells cells = copiedSheet.Cells;
        int maxRow = cells.MaxRow;
        int maxColumn = cells.MaxColumn;

        // If the sheet contains any data, clear the contents of the used range
        if (maxRow >= 0 && maxColumn >= 0)
        {
            cells.ClearContents(0, 0, maxRow, maxColumn);
        }

        // Save the workbook with the duplicated and cleared worksheet
        workbook.Save("output.xlsx");
    }
}
