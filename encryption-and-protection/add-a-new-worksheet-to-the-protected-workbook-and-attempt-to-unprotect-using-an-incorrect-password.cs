// Title: Add a Worksheet to a Structure‑Protected Workbook and Handle Wrong‑Password Unprotect – Aspose.Cells C# Example
// Description: Creates a new Workbook, protects its structure with a password, adds a worksheet named "NewSheet" while the workbook remains protected, attempts to unprotect using an incorrect password, catches the resulting exception, and saves the file with protection still applied.
// Keywords: Aspose.Cells | C# workbook protection | protect workbook structure | add worksheet to protected workbook | unprotect with wrong password | exception handling Aspose.Cells | save protected workbook
// Common Searches: how to add a sheet to a structure‑protected workbook using Aspose.Cells | what error is thrown when unprotecting with an invalid password in Aspose.Cells | C# example for handling wrong password during workbook unprotect | can you add worksheets to a protected workbook in Aspose.Cells | Aspose.Cells unprotect incorrect password exception
// Developer Intent: The developer needs to insert a new worksheet into a workbook whose structure is locked, then verify that calling Unprotect with an invalid password raises an exception and leaves the protection intact.
// Use Cases: Demonstrate that structure protection does not block adding new worksheets. | Show proper try‑catch handling for a failed Unprotect call with a wrong password. | Ensure the workbook is saved while retaining its protection after an unsuccessful unprotect attempt.
// AI Prompts: Generate C# code with Aspose.Cells that adds a sheet to a structure‑protected workbook and safely attempts to unprotect it using an incorrect password, handling any exception. | Explain why Aspose.Cells allows adding worksheets to a workbook with structure protection but throws an error when Unprotect is called with a wrong password. | Create a C# unit test that confirms the workbook remains protected after an unsuccessful Unprotect operation with an invalid password.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a new Workbook, protects its structure with a password, adds a worksheet named "NewSheet" while the workbook remains protected, attempts to unprotect using an incorrect password, catches the resulting exception, and saves the file with protection still applied.
    public class WorkbookUnprotectIncorrectPasswordDemo
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Protect the workbook structure with a password
            workbook.Protect(ProtectionType.Structure, "correctPassword");

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

            // Save the workbook (still protected)
            workbook.Save("ProtectedWorkbook_WithNewSheet.xlsx");
        }
    }
}
