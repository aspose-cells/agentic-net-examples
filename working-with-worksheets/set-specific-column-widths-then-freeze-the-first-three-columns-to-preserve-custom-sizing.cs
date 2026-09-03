// Title: Set column widths and freeze the first three columns in an Excel worksheet with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to assign 20, 30, and 25 character widths to columns A, B, and C respectively, then lock columns A‑C while keeping rows scrollable. | Show how to apply the SetColumnWidth method followed by FreezePanes to freeze only the first three columns in a .NET workbook.
// Common Searches: Aspose.Cells C# how to set column A width to 20 characters | freeze only columns A to C with Aspose.Cells without freezing rows | C# example for setting multiple column widths and then applying FreezePanes | Aspose.Cells SetColumnWidth and FreezePanes usage for custom column sizing
// Tags: Aspose.Cells SetColumnWidth method C# | freeze columns only Aspose.Cells | custom column sizing Excel .NET | FreezePanes for columns Aspose.Cells | Excel worksheet column width configuration C#

using Aspose.Cells;
using System;

// Creates a new workbook, sets column A‑C widths to 20, 30, and 25 characters using SetColumnWidth, freezes columns A‑C with FreezePanes while leaving rows unfrozen, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set specific column widths (in characters)
            // Example widths: Column A = 20, Column B = 30, Column C = 25
            sheet.Cells.SetColumnWidth(0, 20); // Column A
            sheet.Cells.SetColumnWidth(1, 30); // Column B
            sheet.Cells.SetColumnWidth(2, 25); // Column C

            // Freeze the first three columns (A:C) without freezing any rows.
            // totalRows = 0 (no rows frozen), totalColumns = 3 (freeze first three columns)
            // rows = 0, columns = 3 define the top‑left cell of the scrollable area (cell D1)
            sheet.FreezePanes(0, 3, 0, 3);

            // Save the workbook to a file
            workbook.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
