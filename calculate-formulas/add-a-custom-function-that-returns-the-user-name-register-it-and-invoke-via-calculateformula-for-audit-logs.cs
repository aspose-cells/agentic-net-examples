// Title: Create and register a custom user‑name worksheet function and call it with CalculateFormula for audit logging in Aspose.Cells for .NET
// Common Searches: how to define a custom worksheet function that returns Environment.UserName in Aspose.Cells C# | registering a user name UDF and using CalculateFormula for audit log in Aspose.Cells .NET | invoke custom function from a cell formula to log current Windows user in Aspose.Cells workbook | Aspose.Cells CalculateFormula with custom function for audit trail in C# | C# example of adding a custom Excel function that outputs the logged‑in user name
// Tags: custom worksheet function registration Aspose.Cells | CalculateFormula custom user name UDF .NET | audit log workbook using custom function | Environment.UserName Excel UDF C# | Aspose.Cells user name function example

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example demonstrates how to implement a custom worksheet function that returns the current Windows user name via Environment.UserName, register the function with a Workbook using Aspose.Cells' ICustomFunction interface, invoke it through CalculateFormula to obtain the value for audit‑logging, write the result into a cell, output the name to the console, and save the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Retrieve the current Windows user name
                string userName = Environment.UserName;

                // Write the user name into the first worksheet, cell A1
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue(userName);

                // Output the result for audit logging
                Console.WriteLine($"Audit Log - UserName retrieved: {userName}");

                // Save the workbook
                string outputPath = "AuditLogWorkbook.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
