// Title: Aspose.Cells .NET: Set and Remove Write‑Protection Password on an Excel Workbook
// Description: Demonstrates how to apply write‑protection with a password to a new Workbook using Settings.WriteProtection, save the protected file, reload it to verify IsWriteProtected and validate the password, then clear the password to remove protection and save an unprotected copy.
// Keywords: Aspose.Cells write protection | Excel workbook password .NET | remove write protection Aspose.Cells | Settings.WriteProtection.Password | ValidatePassword Aspose.Cells | C# Aspose.Cells example | protect Excel file programmatically
// Common Searches: Aspose.Cells set write protection password C# | How to remove write protection from an Excel workbook using Aspose.Cells | Validate write protection password before clearing Aspose.Cells | C# code to protect and unprotect Excel file with Aspose.Cells | Aspose.Cells write protection IsWriteProtected flag
// Developer Intent: Apply a password to protect an Excel workbook, then programmatically remove the protection using the same password with Aspose.Cells for .NET.
// Use Cases: Create a fresh workbook, assign a write‑protect password, and distribute the protected file. | Load an existing protected workbook, confirm protection status, validate the password, and clear it to allow editing. | Save both protected and unprotected versions and verify that protection has been successfully removed.
// AI Prompts: Generate C# code that sets a write‑protect password on an Aspose.Cells workbook, validates it, and then removes the protection. | Show how to check the IsWriteProtected flag and use ValidatePassword before clearing write protection in Aspose.Cells for .NET. | Provide an example with error handling for applying and removing write protection on an Excel file using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsWriteProtectionDemo
{
    // Demonstrates how to apply write‑protection with a password to a new Workbook using Settings.WriteProtection, save the protected file, reload it to verify IsWriteProtected and validate the password, then clear the password to remove protection and save an unprotected copy.
    class Program
    {
        static void Main()
        {
            // -------------------- Create and protect --------------------
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Apply write protection with a password
            workbook.Settings.WriteProtection.Password = "owner";

            // Save the write‑protected workbook
            string protectedPath = "WriteProtected.xlsx";
            workbook.Save(protectedPath);
            Console.WriteLine($"Workbook saved with write protection: {protectedPath}");

            // -------------------- Load and remove protection --------------------
            // Load the previously saved workbook
            Workbook loadedWorkbook = new Workbook(protectedPath);

            // Verify that the workbook is write protected
            Console.WriteLine("Is write protected (before removal): " +
                loadedWorkbook.Settings.WriteProtection.IsWriteProtected);

            // Validate the password (optional, just to demonstrate)
            bool isValid = loadedWorkbook.Settings.WriteProtection.ValidatePassword("owner");
            Console.WriteLine("Password validation result: " + isValid);

            // Remove write protection by clearing the password
            loadedWorkbook.Settings.WriteProtection.Password = null; // or string.Empty

            // Save the workbook without write protection
            string unprotectedPath = "WriteUnprotected.xlsx";
            loadedWorkbook.Save(unprotectedPath);
            Console.WriteLine($"Workbook saved after removing write protection: {unprotectedPath}");

            // Verify removal
            Workbook verifyWorkbook = new Workbook(unprotectedPath);
            Console.WriteLine("Is write protected (after removal): " +
                verifyWorkbook.Settings.WriteProtection.IsWriteProtected);
        }
    }
}
