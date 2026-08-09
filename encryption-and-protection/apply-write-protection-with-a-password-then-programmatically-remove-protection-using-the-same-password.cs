// Title: C# – Apply and Remove Write‑Protection with Password in Excel using Aspose.Cells
// Description: Demonstrates how to set a write‑protection password on a workbook, optionally protect its structure, verify the protection state, and then programmatically unprotect and clear the password with Aspose.Cells for .NET.
// Keywords: Aspose.Cells write protection | Excel password protection C# | remove workbook protection Aspose.Cells | protect workbook structure .NET | clear write‑protection password | Aspose.Cells unprotect example
// Common Searches: Aspose.Cells set write protection password | C# remove Excel workbook protection with password | how to unprotect workbook structure using Aspose.Cells | clear write‑protection password Aspose.Cells .NET | verify workbook is write protected Aspose.Cells
// Developer Intent: Set a password to write‑protect an Excel workbook and later remove that protection programmatically using the same password.
// Use Cases: Distribute a read‑only template that can be unlocked for authorized updates. | Automate removal of protection in a CI/CD pipeline before further data processing. | Create a secure report for external users, then strip protection for archival.
// AI Prompts: Generate C# code with Aspose.Cells to apply write protection, check IsWriteProtected, and then unprotect using the same password. | Show how to protect workbook structure and later call Unprotect with a password in Aspose.Cells for .NET. | Provide a robust example that catches exceptions when an incorrect password is used to clear write protection.

using System;
using Aspose.Cells;

namespace AsposeCellsWriteProtectionDemo
{
    // Demonstrates how to set a write‑protection password on a workbook, optionally protect its structure, verify the protection state, and then programmatically unprotect and clear the password with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();

            // Add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Write protection demo");

            // ---------- Apply write protection ----------
            // Set the password that restricts modification of the file
            workbook.Settings.WriteProtection.Password = "owner";

            // Additionally protect the workbook structure (optional, shows unprotect usage)
            workbook.Protect(ProtectionType.All, "owner");

            // Save the protected workbook
            string protectedPath = "WriteProtectedWorkbook.xlsx";
            workbook.Save(protectedPath);
            Console.WriteLine($"Workbook saved with write protection: {protectedPath}");

            // ---------- Load the protected workbook ----------
            Workbook loadedWorkbook = new Workbook(protectedPath);

            // Verify that the workbook is write‑protected
            Console.WriteLine("Is write protected? " + loadedWorkbook.Settings.WriteProtection.IsWriteProtected);
            Console.WriteLine("Is structure protected with password? " + loadedWorkbook.IsWorkbookProtectedWithPassword);

            // ---------- Remove protection using the same password ----------
            // Unprotect the workbook structure
            loadedWorkbook.Unprotect("owner");

            // Clear the write‑protection password
            loadedWorkbook.Settings.WriteProtection.Password = null;

            // Save the unprotected workbook
            string unprotectedPath = "UnprotectedWorkbook.xlsx";
            loadedWorkbook.Save(unprotectedPath);
            Console.WriteLine($"Workbook saved after removing protection: {unprotectedPath}");
        }
    }
}
