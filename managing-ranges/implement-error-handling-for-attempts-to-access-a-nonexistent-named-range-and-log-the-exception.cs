// Title: Error Handling for Missing Named Ranges in Aspose.Cells (.NET)
// Description: Demonstrates how to safely retrieve a named range with GetRangeByName, detect a null result, throw and catch an InvalidOperationException, log the error, and protect workbook.Save with a second try‑catch block.
// Keywords: Aspose.Cells | C# | .NET | named range | GetRangeByName | exception handling | error logging | try‑catch | workbook save failure
// Common Searches: Aspose.Cells check if named range exists | catch exception for missing named range C# | log error when GetRangeByName returns null | protect workbook.Save with try catch Aspose.Cells
// Developer Intent: The developer needs a reliable pattern to verify a named range's existence, raise a meaningful exception when it is absent, and record the failure without crashing the application.
// Use Cases: Validate user‑supplied range names before calculations to prevent runtime errors. | Continue automated report generation when a required range has been renamed or deleted, while capturing the issue in logs. | Handle file‑system or permission problems during workbook.Save and log detailed diagnostics.
// AI Prompts: Write a reusable method GetNamedRangeOrThrow that returns a Range or throws a custom MissingRangeException. | Generate code to log exception details (message, stack trace, timestamp) to a file or monitoring system when GetRangeByName fails. | Provide an example of a global error handler that captures both named‑range lookup failures and workbook save errors in an Aspose.Cells application.

using System;
using Aspose.Cells;

// Demonstrates how to safely retrieve a named range with GetRangeByName, detect a null result, throw and catch an InvalidOperationException, log the error, and protect workbook.Save with a second try‑catch block.
class NamedRangeErrorHandling
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Sheet1";

        // Add some sample data
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);

        // Create a valid named range for demonstration
        sheet.Cells.CreateRange("A1:A2").Name = "ValidRange";

        // Attempt to access a non‑existent named range and handle the error
        try
        {
            // GetRangeByName returns null if the named range does not exist
            Aspose.Cells.Range missingRange = workbook.Worksheets.GetRangeByName("MissingRange");
            if (missingRange == null)
            {
                // Throw an exception to be caught below
                throw new InvalidOperationException("Named range 'MissingRange' does not exist.");
            }

            // If the range existed, you could work with it here
            Console.WriteLine("Range address: " + missingRange.Address);
        }
        catch (Exception ex)
        {
            // Log the exception details
            Console.WriteLine("Error accessing named range: " + ex.Message);
        }

        // Save the workbook (lifecycle rule) with safety handling
        try
        {
            workbook.Save("NamedRangeErrorDemo.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error saving workbook: " + ex.Message);
        }
    }
}
