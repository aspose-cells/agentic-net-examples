// Title: Safely applying Worksheet.FreezePanes in C# with Aspose.Cells using try‑catch to manage invalid row or column indices
// AI Prompts: Wrap the Worksheet.FreezePanes call in a try‑catch block that catches ArgumentOutOfRangeException and logs a clear message. | Add pre‑validation for row and column parameters before invoking FreezePanes, and fall back to default behavior when values are out of range. | Implement a helper method that executes FreezePanes, returns a success flag, and provides user‑friendly error details for invalid indices.
// Common Searches: C# Aspose.Cells how to catch ArgumentOutOfRangeException from FreezePanes | example of using try-catch around Worksheet.FreezePanes in .NET | prevent crash when FreezePanes receives negative row index Aspose.Cells | validate freeze pane coordinates before calling FreezePanes in Aspose.Cells | handle invalid freeze pane parameters in Aspose.Cells workbook creation
// Tags: Aspose.Cells FreezePanes exception handling | C# try-catch worksheet freeze panes | validate freeze pane indices Aspose.Cells | out-of-range FreezePanes parameters handling | error handling for worksheet FreezePanes Aspose.Cells | protect FreezePanes call with try-catch

using System;
using Aspose.Cells;

// // Demonstrates creating a workbook, accessing the first worksheet, freezing the first row and column using Worksheet.FreezePanes inside a try‑catch block, and saving the file while gracefully handling any exceptions such as invalid row or column indices.
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

            // Freeze the first row and first column
            // Parameters: row, column, totalRows, totalColumns
            sheet.FreezePanes(1, 1, 1, 1);

            // Save the workbook
            workbook.Save("output.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            // Handle any errors that occur during processing
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
