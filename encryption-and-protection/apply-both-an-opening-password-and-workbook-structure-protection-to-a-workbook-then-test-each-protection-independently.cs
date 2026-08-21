// Title: Apply Opening Password and Structure Protection to an Excel Workbook with Aspose.Cells (.NET) and Verify Independently
// Description: Demonstrates how to create a workbook, set an opening (encryption) password, protect the workbook structure with a separate password, save the file, test each protection separately, unprotect the structure while keeping the opening password, and confirm the remaining encryption on reload.
// Keywords: Aspose.Cells opening password | workbook structure protection .NET | Excel encryption Aspose.Cells | unprotect workbook password | LoadOptions password Aspose.Cells | verify workbook protection
// Common Searches: Aspose.Cells set opening password and protect structure | test Excel file encryption and structure protection with Aspose.Cells | remove workbook structure protection but keep opening password .NET | how to load password‑protected workbook using Aspose.Cells | check if workbook is encrypted Aspose.Cells
// Developer Intent: The developer wants to secure an Excel file with both an opening password and a structure lock, validate each protection independently, and later remove the structure lock while preserving the opening password.
// Use Cases: Create a new workbook, assign an opening password, and protect its structure before saving. | Attempt to open the protected file without a password to confirm encryption enforcement. | Load the file with the correct opening password, inspect encryption and structure flags, and handle incorrect unprotect attempts. | Remove the structure protection using the proper password, save the workbook, and verify that the opening password remains active.
// AI Prompts: Generate C# code with Aspose.Cells that adds an opening password and a separate structure protection to a workbook, then shows how to load the file with the opening password and read protection status. | Provide a C# example that catches exceptions when opening a password‑protected workbook without a password and when unprotecting the structure with an invalid password using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsProtectionDemo
{
    // Demonstrates how to create a workbook, set an opening (encryption) password, protect the workbook structure with a separate password, save the file, test each protection separately, unprotect the structure while keeping the opening password, and confirm the remaining encryption on reload.
    class Program
    {
        static void Main()
        {
            // ------------------- Create and protect workbook -------------------
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Data");

            // Apply an opening (encryption) password
            workbook.Settings.Password = "open123";

            // Protect the workbook structure with a separate password
            workbook.Protect(ProtectionType.Structure, "struct123");

            // Save the protected workbook
            string protectedPath = "ProtectedWorkbook.xlsx";
            workbook.Save(protectedPath);
            workbook.Dispose();

            // ------------------- Test opening password -------------------
            // Attempt to open without password (should throw)
            try
            {
                Workbook wrongLoad = new Workbook(protectedPath);
                wrongLoad.Dispose(); // Not expected to reach here
            }
            catch (Exception ex)
            {
                Console.WriteLine("Opening without password failed as expected: " + ex.Message);
            }

            // Open with the correct opening password
            LoadOptions loadOptions = new LoadOptions { Password = "open123" };
            Workbook loadedWorkbook = new Workbook(protectedPath, loadOptions);

            // Verify encryption and structure protection status
            Console.WriteLine("Is workbook encrypted (requires opening password): " + loadedWorkbook.Settings.IsEncrypted);
            Console.WriteLine("Is workbook protected with password (structure/window): " + loadedWorkbook.IsWorkbookProtectedWithPassword);
            Console.WriteLine("Workbook settings IsProtected (structure/window): " + loadedWorkbook.Settings.IsProtected);

            // ------------------- Test structure protection -------------------
            // Attempt to unprotect with an incorrect password
            try
            {
                loadedWorkbook.Unprotect("wrongPassword");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unprotect with wrong password failed as expected: " + ex.Message);
            }

            // Unprotect with the correct structure password
            loadedWorkbook.Unprotect("struct123");
            Console.WriteLine("After correct unprotect, IsWorkbookProtectedWithPassword: " + loadedWorkbook.IsWorkbookProtectedWithPassword);
            Console.WriteLine("After correct unprotect, Settings.IsProtected: " + loadedWorkbook.Settings.IsProtected);

            // Save the workbook after removing structure protection (opening password remains)
            string unprotectedPath = "UnprotectedWorkbook.xlsx";
            loadedWorkbook.Save(unprotectedPath);
            loadedWorkbook.Dispose();

            // ------------------- Verify that opening password still works -------------------
            // Load the newly saved file with the opening password
            LoadOptions verifyOptions = new LoadOptions { Password = "open123" };
            Workbook verifyWorkbook = new Workbook(unprotectedPath, verifyOptions);
            Console.WriteLine("Verification load - IsEncrypted: " + verifyWorkbook.Settings.IsEncrypted);
            Console.WriteLine("Verification load - IsWorkbookProtectedWithPassword: " + verifyWorkbook.IsWorkbookProtectedWithPassword);
            verifyWorkbook.Dispose();
        }
    }
}
