// Title: Freeze the leftmost two columns in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a workbook with Aspose.Cells, freezes columns A and B, and saves it as an .xlsx file. | Show how to call Worksheet.FreezePanes to lock the first two columns while keeping rows unfrozen in Aspose.Cells. | Provide a step‑by‑step example of using the four‑parameter FreezePanes overload to freeze columns in a .NET Excel file.
// Common Searches: aspnet freeze first two columns excel using aspose.cells | c# Aspose.Cells FreezePanes column index 2 example | how to lock leftmost columns in Excel workbook with Aspose.Cells .NET | freeze columns A B in generated Excel file using Aspose.Cells C#
// Tags: Aspose.Cells FreezePanes column freeze C# | freeze leftmost columns Excel Aspose.Cells | C# Aspose.Cells worksheet FreezePanes overload | generate frozen columns .xlsx with Aspose.Cells | Excel column locking using Aspose.Cells .NET

using Aspose.Cells;
using System;

// The example creates a new Workbook, accesses the first Worksheet, calls sheet.FreezePanes(0, 2, 0, 0) to freeze columns A and B while leaving rows unfrozen, and saves the result as FrozenColumns.xlsx, with basic exception handling.
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

            // Freeze the leftmost two columns (row index = 0, column index = 2)
            // Use the overload with four parameters; totalRows and totalColumns set to 0
            sheet.FreezePanes(0, 2, 0, 0);

            // Save the workbook to a file
            workbook.Save("FrozenColumns.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
