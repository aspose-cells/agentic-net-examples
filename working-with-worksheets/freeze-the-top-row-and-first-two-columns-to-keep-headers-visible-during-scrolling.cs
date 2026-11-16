// Title: How to freeze the first row and the first two columns in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# freeze top row and first two columns in Excel | C# code to set frozen panes for header row and columns with Aspose.Cells | Freeze panes on worksheet using Aspose.Cells .NET API | How to keep Excel headers visible while scrolling with Aspose.Cells | Programmatically freeze rows and columns in an Excel file using Aspose.Cells C#
// Tags: Aspose.Cells FreezePanes C# | freeze top row Excel Aspose.Cells | freeze first two columns Excel Aspose | Excel worksheet frozen panes .NET | programmatic frozen panes Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a new Workbook, accesses the first Worksheet, applies FreezePanes(1, 2, 1, 2) to lock the top row and the first two columns, and saves the file as FrozenPane.xlsx, with exception handling for any runtime errors.
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

            // Freeze the top row and the first two columns.
            // Parameters: row, column, totalRows, totalColumns
            // row = 1 (freeze rows above row 1), column = 2 (freeze columns left of column 2)
            // totalRows = 1 (freeze one row), totalColumns = 2 (freeze two columns)
            sheet.FreezePanes(1, 2, 1, 2);

            // Save the workbook
            workbook.Save("FrozenPane.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
