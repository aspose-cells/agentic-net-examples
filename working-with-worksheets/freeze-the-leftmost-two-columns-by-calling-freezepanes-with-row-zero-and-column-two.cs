// Title: Freeze the First Two Columns (A & B) in Excel Using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, access the first worksheet, and call FreezePanes(0, 2, 0, 2) to lock columns A and B while keeping all rows scrollable, then saves the file as FreezeLeftTwoColumns.xlsx.
// Keywords: Aspose.Cells FreezePanes C# | freeze first two columns Excel | lock columns A B Aspose.Cells | C# Excel freeze panes example | Aspose.Cells .NET freeze columns | Excel worksheet freeze columns programmatically
// Common Searches: Aspose.Cells freeze first two columns C# | How to lock columns A and B in Excel with Aspose.Cells | FreezePanes method example .NET | C# code to freeze columns without rows in Excel | Aspose.Cells freeze panes tutorial
// Developer Intent: The developer needs to freeze the leftmost two columns of a worksheet while leaving all rows unfrozen.
// Use Cases: Create a template where identifier columns stay visible during horizontal scrolling. | Generate reports that require static reference columns (e.g., IDs, names) while data scrolls. | Prepare exported Excel files for end‑users who need the first two columns always in view.
// AI Prompts: Show how to freeze the first three columns in an Excel sheet using Aspose.Cells FreezePanes (C#). | Provide a C# example that freezes both the top row and the first two columns with Aspose.Cells. | Explain each parameter of the FreezePanes method and how to unfreeze panes later in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, access the first worksheet, and call FreezePanes(0, 2, 0, 2) to lock columns A and B while keeping all rows scrollable, then saves the file as FreezeLeftTwoColumns.xlsx.
class FreezeLeftTwoColumns
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Freeze the leftmost two columns (A and B)
        // row = 0 (no rows frozen), column = 2 (freeze at column index 2, i.e., after B)
        // freezedRows = 0, freezedColumns = 2 (freeze two columns)
        worksheet.FreezePanes(0, 2, 0, 2);

        // Save the workbook
        workbook.Save("FreezeLeftTwoColumns.xlsx");
    }
}
