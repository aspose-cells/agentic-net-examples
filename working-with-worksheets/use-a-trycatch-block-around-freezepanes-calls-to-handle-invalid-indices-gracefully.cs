// Title: C# – Wrap Aspose.Cells Worksheet.FreezePanes in Try‑Catch to Safely Handle Invalid Indices
// Description: Demonstrates how to create a Workbook, access the first Worksheet, and call Worksheet.FreezePanes twice—once with a valid cell reference (C3) and three frozen rows/columns, and once with invalid zero indices. Each call is enclosed in its own try‑catch block to capture exceptions, log a friendly message, and allow the program to continue before saving the file as FreezePanesDemo.xlsx.
// Keywords: Aspose.Cells | Worksheet.FreezePanes | C# | .NET | exception handling | try catch | invalid indices | freeze panes error handling | Aspose.Cells example
// Common Searches: Aspose.Cells FreezePanes try catch example | how to handle invalid FreezePanes parameters in C# | catch exception Worksheet.FreezePanes Aspose.Cells | freeze panes zero row column error Aspose.Cells
// Developer Intent: Show how to protect Worksheet.FreezePanes calls with try‑catch blocks so that out‑of‑range or zero indices do not crash the application.
// Use Cases: Freeze panes at a specific cell (e.g., C3) while ensuring runtime errors are caught. | Attempt a FreezePanes operation with invalid parameters, capture the exception, and display a custom message. | Log FreezePanes failures, continue processing other workbook tasks, and still save the final file.
// AI Prompts: Generate C# code using Aspose.Cells that freezes panes at a given cell and includes try‑catch for invalid arguments. | Explain which Worksheet.FreezePanes overloads throw exceptions for out‑of‑range indices and how to handle them. | Provide best practices for logging FreezePanes errors in an Aspose.Cells workflow.

using System;
using Aspose.Cells;

namespace AsposeCellsFreezePanesDemo
{
    // Demonstrates how to create a Workbook, access the first Worksheet, and call Worksheet.FreezePanes twice—once with a valid cell reference (C3) and three frozen rows/columns, and once with invalid zero indices. Each call is enclosed in its own try‑catch block to capture exceptions, log a friendly message, and allow the program to continue before saving the file as FreezePanesDemo.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Example of a valid FreezePanes call
            try
            {
                // Freeze panes at cell C3 with 3 frozen rows and 3 frozen columns
                worksheet.FreezePanes("C3", 3, 3);
                Console.WriteLine("Valid FreezePanes executed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during valid FreezePanes: {ex.Message}");
            }

            // Example of an invalid FreezePanes call that will cause an exception
            try
            {
                // Attempt to freeze panes with invalid indices (row, column, frozen rows, frozen columns all zero)
                worksheet.FreezePanes(0, 0, 0, 0);
            }
            catch (Exception ex)
            {
                // Gracefully handle the error
                Console.WriteLine($"Handled invalid FreezePanes call: {ex.Message}");
            }

            // Save the workbook (save rule)
            workbook.Save("FreezePanesDemo.xlsx");
        }
    }
}
