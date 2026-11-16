// Title: Using try‑catch to safely call Worksheet.FreezePanes in Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells how to catch exception when FreezePanes indices are out of range | example of wrapping Worksheet.FreezePanes in try catch to avoid runtime error | Aspose.Cells FreezePanes invalid row or column index error handling in .NET | prevent crash by handling FreezePanes exceptions in C# workbook code | best practice for exception handling around FreezePanes method Aspose.Cells
// Tags: Worksheet FreezePanes error handling | Aspose.Cells try-catch usage | invalid FreezePanes indices handling | C# freeze panes exception management | protect workbook freeze operation

using System;
using Aspose.Cells;

// The sample creates a Workbook, fills the first worksheet with data, and attempts to freeze the first row and column using Worksheet.FreezePanes inside a try‑catch block that logs any exception, then saves the file with additional error handling.
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

            // Populate sample data
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["B2"].PutValue(456);

            // Freeze the first row and first column (row index 1, column index 1)
            // totalRows = 1, totalColumns = 1 to freeze that row/column
            try
            {
                sheet.FreezePanes(1, 1, 1, 1);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to freeze panes: {ex.Message}");
            }

            // Save the workbook
            try
            {
                workbook.Save("Output.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
