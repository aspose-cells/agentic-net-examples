// Title: Add a Sheet to a Password‑Protected Workbook and Handle Wrong‑Password Unprotect in Aspose.Cells for .NET
// Description: Creates a workbook, applies password protection, adds a new worksheet while the workbook stays protected, attempts to unprotect with an incorrect password, catches the resulting exception, checks protection status via IsWorkbookProtectedWithPassword, and saves the still‑protected file.
// Keywords: Aspose.Cells | C# workbook protection | protect workbook with password | add worksheet to protected workbook | unprotect workbook wrong password | IsWorkbookProtectedWithPassword | exception handling Aspose.Cells | save protected workbook | .NET Excel security
// Common Searches: How to add a sheet to a password protected Excel file using Aspose.Cells C# | What happens when workbook.Unprotect is called with an incorrect password in Aspose.Cells | Check if a workbook is still protected after a failed unprotect attempt | Aspose.Cells C# example for protecting and unprotecting a workbook | Exception thrown by workbook.Unprotect with wrong password
// Developer Intent: Add a worksheet to a password‑protected workbook and verify that an incorrect unprotect attempt fails without removing protection.
// Use Cases: Demonstrate that protection remains active when new sheets are added to a secured workbook. | Show proper exception handling for workbook.Unprotect when the supplied password is invalid. | Validate protection status after a failed unprotect operation using IsWorkbookProtectedWithPassword.
// AI Prompts: Generate C# code with Aspose.Cells that protects a workbook, adds a new worksheet, tries to unprotect using a wrong password, catches the exception, and confirms the workbook stays protected. | Explain the behavior of Aspose.Cells when workbook.Unprotect receives an incorrect password and how IsWorkbookProtectedWithPassword reflects the outcome. | Provide step‑by‑step guidance for safely inserting sheets into a password‑protected workbook without disabling its protection in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, applies password protection, adds a new worksheet while the workbook stays protected, attempts to unprotect with an incorrect password, catches the resulting exception, checks protection status via IsWorkbookProtectedWithPassword, and saves the still‑protected file.
    class WorkbookUnprotectDemo
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Protect the entire workbook with a password
            workbook.Protect(ProtectionType.All, "correctPassword");

            // Add a new worksheet to the protected workbook
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
            Console.WriteLine("Workbook is still protected: " + workbook.IsWorkbookProtectedWithPassword);

            // Save the workbook (still protected)
            workbook.Save("ProtectedWorkbook.xlsx");
        }
    }
}
