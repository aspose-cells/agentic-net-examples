// Title: Set column widths and freeze the first three columns in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set column width in characters | Freeze first three columns in Excel with Aspose.Cells .NET | C# example to apply custom column widths before freezing panes using Aspose.Cells | How to use FreezePanes to lock columns A to C in Aspose.Cells | Set column width for multiple columns and then freeze panes in an Aspose.Cells workbook
// Tags: set column width Aspose.Cells C# | freeze panes first three columns Aspose.Cells | custom column sizing Excel Aspose.Cells | Aspose.Cells FreezePanes usage | column width by character count Aspose.Cells

using Aspose.Cells;
using System;

// The sample creates a new Workbook, accesses the first Worksheet, sets column widths for columns A, B, and C to 20, 25, and 15 characters respectively, freezes columns A‑C using the FreezePanes method, and saves the workbook as Output.xlsx while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set specific column widths (in characters)
            sheet.Cells.SetColumnWidth(0, 20); // Column A
            sheet.Cells.SetColumnWidth(1, 25); // Column B
            sheet.Cells.SetColumnWidth(2, 15); // Column C
            // Add more column width settings as needed
            // sheet.Cells.SetColumnWidth(3, 30); // Example for Column D

            // Freeze the first three columns (A, B, C)
            // Parameters: totalRows, totalColumns, frozenRows, frozenColumns
            sheet.FreezePanes(0, 3, 0, 3);

            // Save the workbook
            workbook.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
