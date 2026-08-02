// Title: C# – Freeze the Top 3 Rows in an Excel Worksheet with Aspose.Cells
// Description: Demonstrates how to lock the first three rows of a worksheet using Aspose.Cells for .NET. The example creates a workbook, applies FreezePanes(3,0,3,0) so rows 1‑3 stay visible while scrolling, adds sample data, and saves the file as FreezeTopThreeRows.xlsx.
// Keywords: Aspose.Cells FreezePanes C# | freeze top rows Excel .NET | lock header rows Aspose.Cells | freeze first three rows C# example | Excel freeze panes programmatically
// Common Searches: how to freeze the first three rows in Excel using Aspose.Cells | Aspose.Cells C# freeze top rows without freezing columns | FreezePanes parameters for locking header rows in .NET | C# code to keep top rows visible in an Excel file
// Developer Intent: The developer needs to keep the initial three rows of a worksheet visible while scrolling through the rest of the data.
// Use Cases: Financial statements where summary rows at the top must remain in view. | Exported reports with multi‑level column headers occupying the first three rows. | Dashboard templates that require a fixed title block and subtitle rows.
// AI Prompts: Generate C# code that freezes a configurable number of top rows in an Excel sheet using Aspose.Cells. | Explain each argument of the FreezePanes method and show how to freeze rows only, leaving columns scrollable. | Provide a combined example that freezes both rows and columns in a worksheet with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Demonstrates how to lock the first three rows of a worksheet using Aspose.Cells for .NET. The example creates a workbook, applies FreezePanes(3,0,3,0) so rows 1‑3 stay visible while scrolling, adds sample data, and saves the file as FreezeTopThreeRows.xlsx.
class FreezeTopRowsExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Freeze the top three rows (rows 0,1,2). 
        // Freeze at row index 3 (the first unfrozen row) and column index 0.
        worksheet.FreezePanes(3, 0, 3, 0);

        // Add sample data to visualize the frozen rows
        for (int i = 0; i < 10; i++)
        {
            worksheet.Cells[i, 0].PutValue($"Row {i + 1}");
        }

        // Save the workbook
        workbook.Save("FreezeTopThreeRows.xlsx");
    }
}
