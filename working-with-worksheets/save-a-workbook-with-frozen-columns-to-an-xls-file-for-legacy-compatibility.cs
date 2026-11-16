// Title: Create an Excel 97‑2003 (.xls) file with frozen columns using Aspose.Cells for .NET (C#)
// Common Searches: how to freeze the first two columns and save as .xls using Aspose.Cells in C# | Aspose.Cells C# freeze panes then export to Excel 97-2003 format | save workbook with frozen columns to legacy XLS file with Aspose.Cells | C# code to generate an XLS file that has column freeze using Aspose.Cells | freeze columns in Aspose.Cells and output an Excel 97‑2003 workbook
// Tags: freeze panes Aspose.Cells C# | export to Excel97To2003 format Aspose.Cells | legacy XLS creation with column freeze | Aspose.Cells column freeze example | XLS output with frozen columns Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new Workbook, freezes the first two columns (A and B) on the first worksheet with FreezePanes, optionally writes header values, and saves the file as a legacy XLS workbook by using SaveFormat.Excel97To2003, all wrapped in a try‑catch block.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Freeze the first two columns (A and B)
            // row = 0 (no frozen rows), column = 2 (first unfrozen column is C)
            // freezedRows = 0, freezedColumns = 2 (freeze two columns)
            sheet.FreezePanes(0, 2, 0, 2);

            // Example data (optional)
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["C1"].PutValue("Header3");

            // Save the workbook as an XLS file for legacy compatibility
            workbook.Save("FrozenColumns.xls", SaveFormat.Excel97To2003);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
