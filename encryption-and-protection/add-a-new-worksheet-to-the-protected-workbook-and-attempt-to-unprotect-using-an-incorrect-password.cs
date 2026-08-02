// Title: Add a Worksheet to a Password‑Protected Workbook and Handle an Incorrect Unprotect Password – Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, protect it with a password, add a new sheet while protection is active, attempt to unprotect using a wrong password (capturing the exception), verify that the workbook remains protected, and save the file.
// Keywords: Aspose.Cells protect workbook .NET | add worksheet to protected workbook | unprotect workbook wrong password | Workbook.Settings.IsProtected check | save protected workbook C# | encryption and protection Aspose.Cells | Aspose.Cells .NET US developers | Aspose.Cells Europe examples
// Common Searches: How to add a sheet to a password‑protected workbook with Aspose.Cells | What error is thrown when Unprotect is called with an incorrect password | Verify workbook protection after a failed unprotect attempt | Saving a protected workbook after modifying worksheets in C#
// Developer Intent: Show how to insert a new worksheet into an already password‑protected workbook and confirm that an invalid unprotect call does not remove the protection.
// Use Cases: Maintain workbook security while programmatically adding additional worksheets. | Gracefully handle wrong password errors during unprotect operations. | Ensure the protection flag stays true after a failed unprotect attempt before saving.
// AI Prompts: Generate C# code using Aspose.Cells that protects a workbook, adds a new worksheet, tries to unprotect with an incorrect password, catches the exception, and confirms the workbook remains protected. | Create a unit test in C# that asserts Aspose.Cells throws an exception for a wrong unprotect password and that Settings.IsProtected stays true after the attempt.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, protect it with a password, add a new sheet while protection is active, attempt to unprotect using a wrong password (capturing the exception), verify that the workbook remains protected, and save the file.
    public class WorkbookAddWorksheetAndWrongUnprotectDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("An unexpected error occurred: " + ex.Message);
            }
        }

        public static void Run()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Protect the entire workbook with a password
            workbook.Protect(ProtectionType.All, "correctPassword");

            // Add a new worksheet to the already protected workbook
            Worksheet newSheet = workbook.Worksheets.Add("NewSheet");

            // Attempt to unprotect the workbook using an incorrect password
            try
            {
                workbook.Unprotect("wrongPassword");
                Console.WriteLine("Workbook unprotected with wrong password (unexpected).");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to unprotect workbook with wrong password: " + ex.Message);
            }

            // Verify that the workbook is still protected
            Console.WriteLine("Workbook is still protected: " + workbook.Settings.IsProtected);

            // Save the workbook (lifecycle: save)
            try
            {
                workbook.Save("ProtectedWorkbook_WithNewSheet.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to save workbook: " + ex.Message);
            }
        }
    }
}
