// Title: Set and Remove Write‑Protection Password on an Excel Workbook with Aspose.Cells (.NET)
// Description: Demonstrates how to apply a write‑protection password to a new Workbook, save it, reload the file, validate the password with ValidatePassword, clear the password to remove protection, and save the unprotected workbook using Aspose.Cells for C#.
// Keywords: Aspose.Cells write protection | Excel write protection password C# | remove write protection Aspose.Cells | ValidatePassword Aspose.Cells | unprotect Excel file programmatically | Aspose.Cells workbook security | C# Excel password protection
// Common Searches: how to add write protection password to Excel using Aspose.Cells | remove write protection from Aspose.Cells workbook | Aspose.Cells validate write protection password before unprotecting | C# code to clear write protection in Excel with Aspose.Cells | Aspose.Cells set and clear workbook write protection
// Developer Intent: Apply a password‑protected write lock to an Excel workbook and later remove it programmatically with the same password.
// Use Cases: Create a template that is locked for editing, distribute it, then unlock it in an automated process to populate data. | Secure a generated report before sending it to a client and later remove protection for further revisions. | Ensure only authorized code can modify a workbook by validating the password before clearing write protection.
// AI Prompts: Show C# code that sets a write‑protection password on an Aspose.Cells Workbook and verifies the protection before saving. | Provide an example of loading a write‑protected Excel file with Aspose.Cells, validating the password, removing the protection, and saving the unprotected file. | Explain how to handle an incorrect password when attempting to remove write protection using Aspose.Cells in C#.

using System;
using Aspose.Cells;

namespace AsposeCellsWriteProtectionDemo
{
    // Demonstrates how to apply a write‑protection password to a new Workbook, save it, reload the file, validate the password with ValidatePassword, clear the password to remove protection, and save the unprotected workbook using Aspose.Cells for C#.
    class Program
    {
        static void Main()
        {
            // -------------------- Create and apply write protection --------------------
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set write protection password
            string writePassword = "owner";
            workbook.Settings.WriteProtection.Password = writePassword;

            // Verify that the workbook is write protected
            Console.WriteLine("Is write protected (before save): " + workbook.Settings.WriteProtection.IsWriteProtected);

            // Save the write‑protected workbook
            string protectedPath = "WriteProtected.xlsx";
            workbook.Save(protectedPath);
            Console.WriteLine("Workbook saved with write protection: " + protectedPath);

            // -------------------- Load and remove write protection --------------------
            // Load the previously saved workbook (no load password required for write protection)
            Workbook loadedWorkbook = new Workbook(protectedPath);

            // Validate the password before attempting to remove protection
            bool isValid = loadedWorkbook.Settings.WriteProtection.ValidatePassword(writePassword);
            Console.WriteLine("Password validation result: " + isValid);

            if (isValid)
            {
                // Remove write protection by clearing the password
                loadedWorkbook.Settings.WriteProtection.Password = null;

                // Verify that protection has been removed
                Console.WriteLine("Is write protected (after removal): " + loadedWorkbook.Settings.WriteProtection.IsWriteProtected);

                // Save the unprotected workbook
                string unprotectedPath = "WriteUnprotected.xlsx";
                loadedWorkbook.Save(unprotectedPath);
                Console.WriteLine("Workbook saved after removing write protection: " + unprotectedPath);
            }
            else
            {
                Console.WriteLine("Incorrect password. Write protection not removed.");
            }
        }
    }
}
