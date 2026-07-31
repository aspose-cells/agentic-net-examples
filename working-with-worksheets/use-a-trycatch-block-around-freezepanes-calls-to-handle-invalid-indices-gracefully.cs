// Title: C# Try‑Catch for Worksheet.FreezePanes – Handle Invalid Indices in Aspose.Cells
// Description: Demonstrates creating a Workbook, catching the exception thrown by an invalid Worksheet.FreezePanes(0,0,0,0) call, then applying a valid FreezePanes(3,3,3,3) inside a separate try‑catch block, and finally saving the file as FreezePanesResult.xlsx.
// Keywords: Aspose.Cells | C# | .NET | Worksheet.FreezePanes | exception handling | try‑catch | invalid indices | freeze panes example | error logging | FreezePanes validation
// Common Searches: Aspose.Cells FreezePanes try catch example | how to catch FreezePanes exception in C# | valid parameters for Worksheet.FreezePanes | handle invalid FreezePanes indices Aspose.Cells | C# code to log FreezePanes errors
// Developer Intent: Show how to wrap Worksheet.FreezePanes calls in try‑catch blocks to gracefully manage parameter‑validation errors.
// Use Cases: Validate user‑supplied row/column values before freezing panes and log any failures. | Prevent application crashes when dynamic data produces out‑of‑range FreezePanes arguments. | Implement fallback freeze settings after an exception occurs during workbook generation.
// AI Prompts: Generate C# code that checks FreezePanes arguments and uses try‑catch to log exceptions in Aspose.Cells. | Provide an Aspose.Cells snippet that records detailed error information when FreezePanes fails. | Explain best practices for handling Worksheet.FreezePanes errors in a .NET reporting tool.

using System;
using Aspose.Cells;

// Demonstrates creating a Workbook, catching the exception thrown by an invalid Worksheet.FreezePanes(0,0,0,0) call, then applying a valid FreezePanes(3,3,3,3) inside a separate try‑catch block, and finally saving the file as FreezePanesResult.xlsx.
class FreezePanesDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Attempt to freeze panes with invalid indices (row, column, frozen rows, frozen columns all zero)
        // This will throw an exception according to Aspose.Cells validation rules.
        try
        {
            sheet.FreezePanes(0, 0, 0, 0);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Caught invalid FreezePanes parameters: " + ex.Message);
        }

        // Perform a valid FreezePanes operation
        try
        {
            // Freeze at cell C3 (row index 3, column index 3) with 3 frozen rows and 3 frozen columns
            sheet.FreezePanes(3, 3, 3, 3);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error during valid FreezePanes call: " + ex.Message);
        }

        // Save the workbook to disk
        workbook.Save("FreezePanesResult.xlsx");
    }
}
