// Title: C# – Catch Missing Named Range Exception and Log It with Aspose.Cells
// Description: Shows how to create a workbook, add a valid named range, then safely try to retrieve a non‑existent named range using Aspose.Cells for .NET. The access is wrapped in a try‑catch block, the exception is logged, and the workbook is saved without terminating the application.
// Keywords: Aspose.Cells | C# | .NET | named range | exception handling | GetRange | null reference | error logging | workbook save | missing range
// Common Searches: Aspose.Cells catch exception for missing named range | C# log error when named range not found | GetRange null reference Aspose.Cells .NET | how to handle non‑existent named range in Aspose.Cells | Aspose.Cells error handling example
// Developer Intent: Add robust try‑catch logic around named‑range access to prevent crashes and record the error details.
// Use Cases: Validate user‑supplied named ranges before processing data. | Prevent runtime failures when imported spreadsheets lack expected ranges. | Capture detailed error information for troubleshooting automated report pipelines.
// AI Prompts: Generate C# code that checks for a named range's existence with Aspose.Cells and writes a warning to the console instead of throwing. | Provide an example of logging the stack trace of a missing named‑range exception to a file using Aspose.Cells. | Suggest best practices for handling null Name objects when calling GetRange in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsNamedRangeErrorHandling
{
    // Shows how to create a workbook, add a valid named range, then safely try to retrieve a non‑existent named range using Aspose.Cells for .NET. The access is wrapped in a try‑catch block, the exception is logged, and the workbook is saved without terminating the application.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Sheet1";

                // Add a valid named range for demonstration
                int validIndex = workbook.Worksheets.Names.Add("ValidRange");
                workbook.Worksheets.Names[validIndex].RefersTo = "=Sheet1!$A$1:$A$3";

                // Attempt to access a non‑existent named range
                try
                {
                    // This will return null because the name does not exist
                    Name missingName = workbook.Worksheets.Names["MissingRange"];

                    // Trying to call GetRange on a null reference throws an exception
                    Aspose.Cells.Range missingRange = missingName.GetRange(); // <-- will throw

                    // If no exception, display the address (unlikely)
                    Console.WriteLine("Missing range address: " + missingRange.Address);
                }
                catch (Exception ex)
                {
                    // Log the exception details
                    Console.WriteLine("Error accessing named range: " + ex.Message);
                }

                // Save the workbook (lifecycle rule)
                workbook.Save("NamedRangeErrorHandling.xlsx");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors during workbook processing
                Console.WriteLine("Unexpected error: " + ex.Message);
            }
        }
    }
}
