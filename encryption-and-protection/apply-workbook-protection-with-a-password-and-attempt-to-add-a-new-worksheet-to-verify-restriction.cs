// Title: C# – Password‑protect workbook structure and block worksheet addition with Aspose.Cells
// Description: Shows how to create a Workbook, lock its structure using a password, try to insert a new sheet (causing an exception), save the file, reload it, and verify the lock with the IsWorkbookProtectedWithPassword flag.
// Keywords: Aspose.Cells | C# | Workbook.Protect | ProtectionType.Structure | password protection | prevent adding worksheets | IsWorkbookProtectedWithPassword | Excel workbook security .NET | structure lock | exception handling for protected workbook
// Common Searches: How to lock Excel workbook structure with a password in C# | Aspose.Cells prevent adding new sheets after protection | Check if a saved workbook is password‑protected using Aspose.Cells | Exception when adding worksheet to a protected workbook | C# code sample for workbook structure protection
// Developer Intent: Apply a password to the workbook’s structure and confirm that inserting additional worksheets is disallowed.
// Use Cases: Distribute an Excel file where users cannot modify the sheet layout. | Automated tests that need to ensure protection settings survive a save‑load cycle. | Graceful handling of protection‑related errors in enterprise reporting tools.
// AI Prompts: Write C# code that uses Aspose.Cells to protect a workbook’s structure with a password, attempts to add a worksheet, and logs the expected error. | Generate a snippet that loads a saved Excel file and programmatically checks the IsWorkbookProtectedWithPassword property. | Explain how to protect only the workbook structure (leaving windows unprotected) and how to catch protection exceptions in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsWorkbookProtectionDemo
{
    // Shows how to create a Workbook, lock its structure using a password, try to insert a new sheet (causing an exception), save the file, reload it, and verify the lock with the IsWorkbookProtectedWithPassword flag.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Protect the workbook structure with a password
            // This prevents adding, deleting, or renaming worksheets
            workbook.Protect(ProtectionType.Structure, "MySecretPwd");

            // Attempt to add a new worksheet – should fail because the structure is protected
            try
            {
                workbook.Worksheets.Add();
                Console.WriteLine("New worksheet added (unexpected).");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unable to add worksheet as expected: " + ex.Message);
            }

            // Save the protected workbook
            string filePath = "ProtectedWorkbook.xlsx";
            workbook.Save(filePath);
            workbook.Dispose();

            // Load the saved workbook to verify protection status
            Workbook loadedWorkbook = new Workbook(filePath);
            Console.WriteLine("IsWorkbookProtectedWithPassword: " + loadedWorkbook.IsWorkbookProtectedWithPassword);
            loadedWorkbook.Dispose();
        }
    }
}
