// Title: Freeze the top three rows in Excel using Aspose.Cells for .NET (C#)
// Description: C# example that creates a workbook, freezes the first three rows (row index 3) while leaving columns unfrozen, and saves the file as FreezeTopThreeRows.xlsx.
// Keywords: Aspose.Cells FreezePanes C# | freeze top rows Excel .NET | freeze first three rows Aspose | C# Excel freeze rows without columns | Aspose.Cells workbook freeze panes
// Common Searches: How to freeze the first three rows in an Excel file with Aspose.Cells | Aspose.Cells FreezePanes method example for rows only | C# code to lock header rows in Excel using Aspose | Freeze rows but not columns in Aspose.Cells for .NET | Save Excel workbook after freezing top rows with Aspose
// Developer Intent: Programmatically lock the top three rows of a worksheet while keeping all columns scrollable.
// Use Cases: Generate reports where header rows stay visible during vertical scrolling. | Create data‑entry templates with fixed title and instruction rows. | Automate Excel exports that require frozen header rows for better readability.
// AI Prompts: Write a C# function that accepts a row count parameter and uses Aspose.Cells to freeze that many top rows. | Explain each parameter of the FreezePanes method and show how to unfreeze panes later. | Demonstrate freezing both rows and columns together with Aspose.Cells and then exporting the workbook to PDF.

using System;
using Aspose.Cells;

namespace FreezeTopRowsDemo
{
    // C# example that creates a workbook, freezes the first three rows (row index 3) while leaving columns unfrozen, and saves the file as FreezeTopThreeRows.xlsx.
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
