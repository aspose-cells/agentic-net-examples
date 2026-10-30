// Title: How to Register and Use a Custom UserName Function with CalculateFormula for Audit Logging in Aspose.Cells for .NET
// Common Searches: register a custom worksheet function that returns the Windows user name in Aspose.Cells C# | use CalculateFormula to call a user-defined function for audit log in Aspose.Cells .NET | example of adding a custom function to Aspose.Cells workbook for logging current user | invoke custom user name function via CalculateFormula in Aspose.Cells C# tutorial | Aspose.Cells calculateformula custom function for audit trail implementation
// Tags: custom worksheet function registration Aspose.Cells | calculateformula custom function .NET | audit log workbook generation Aspose.Cells | environment.username retrieval Aspose.Cells | c# excel audit trail using custom function

using System;
using Aspose.Cells;

namespace AuditLogExample
{
    // The example shows how to create a custom worksheet function that returns the current Windows user name, register it with the workbook, and invoke it through CalculateFormula to populate an audit log cell. It demonstrates creating a Workbook, defining the GetUserName function (which uses Environment.UserName), adding the function to the workbook's CustomFunctions collection, calculating the formula "=GetUserName()", writing the result to a cell, and saving the workbook as an audit log file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook instance
                Workbook workbook = new Workbook();

                // Directly obtain the current Windows user name
                string userName = Environment.UserName;

                // Write the result to a cell for audit logging
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("User:");
                sheet.Cells["B1"].PutValue(userName);

                // Save the workbook containing the audit log
                string outputPath = "AuditLog.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Audit log saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
