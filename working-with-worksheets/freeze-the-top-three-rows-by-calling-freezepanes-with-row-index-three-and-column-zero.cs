// Title: Freeze Top Three Rows in Excel with Aspose.Cells for .NET (C#)
// Description: This example creates a new Workbook, accesses the first Worksheet, and uses Worksheet.FreezePanes(3, 0, 3, 0) to lock the first three rows while leaving columns scrollable, then saves the file as FreezeTopThreeRows.xlsx.
// Keywords: Aspose.Cells FreezePanes C# | freeze first three rows Excel | lock header rows Aspose.Cells | C# Excel freeze panes | Aspose.Cells workbook freeze rows | Excel .NET freeze top rows | FreezePanes method parameters
// Common Searches: How to freeze first three rows using Aspose.Cells C# | Aspose.Cells FreezePanes example .NET | Freeze rows only Aspose.Cells | C# code to freeze top rows in Excel | Aspose.Cells freeze panes without columns
// Developer Intent: Apply FreezePanes to lock the first three rows of a worksheet while keeping all columns unfrozen.
// Use Cases: Financial reports where header rows must stay visible during scrolling. | Data‑entry templates that require static top rows for column titles. | Large data sets where the first three rows contain summary information.
// AI Prompts: Write C# code using Aspose.Cells to freeze the first three rows of a worksheet and save the workbook as an .xlsx file. | Explain the four parameters of Worksheet.FreezePanes(row, column, totalRows, totalColumns) and how to use them to freeze rows only. | Show how to iterate over all worksheets in a workbook and apply the same three‑row freeze setting to each.

using System;
using Aspose.Cells;

namespace AsposeCellsFreezeTopRows
{
    // This example creates a new Workbook, accesses the first Worksheet, and uses Worksheet.FreezePanes(3, 0, 3, 0) to lock the first three rows while leaving columns scrollable, then saves the file as FreezeTopThreeRows.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Freeze the top three rows (row index 3) and no columns (column index 0)
            // Parameters: row index, column index, number of frozen rows, number of frozen columns
            worksheet.FreezePanes(3, 0, 3, 0);

            // Save the workbook to a file
            workbook.Save("FreezeTopThreeRows.xlsx");
        }
    }
}
