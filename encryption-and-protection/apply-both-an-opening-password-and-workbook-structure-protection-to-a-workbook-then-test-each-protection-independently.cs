// Title: Set opening password and workbook structure protection, then verify each with Aspose.Cells for .NET
// Description: This C# example shows how to create an Excel workbook, assign an opening (encryption) password, protect the workbook structure with a separate password, save the file, attempt to open it without a password, open it with the correct password using LoadOptions, check encryption and structure‑protection flags, handle wrong unprotect passwords, and finally remove the structure protection.
// Keywords: Aspose.Cells | C# | .NET | Excel opening password | workbook structure protection | LoadOptions password | IsEncrypted | IsWorkbookProtectedWithPassword | Unprotect workbook | encrypted Excel file | protect workbook programmatically
// Common Searches: Aspose.Cells set opening password | protect workbook structure Aspose.Cells | open encrypted Excel with Aspose.Cells .NET | check workbook protection status Aspose.Cells | remove workbook structure protection C#
// Developer Intent: Demonstrate how to apply both an opening (encryption) password and a workbook‑structure protection password to an Excel file using Aspose.Cells for .NET, and how to test each protection independently.
// Use Cases: Generate a confidential Excel report that requires a password to open and prevents sheet reordering. | Validate that an incoming workbook is encrypted before processing its contents. | Programmatically unlock a workbook’s structure after authenticating a user. | Automate compliance checks for Excel file protection in enterprise workflows.
// AI Prompts: Write C# code with Aspose.Cells to add an opening password and a separate structure‑protection password, then show how to open the file with LoadOptions and handle incorrect passwords. | Explain step‑by‑step how to test encryption and structure protection in Aspose.Cells, including exception handling for wrong passwords. | Recommend secure practices for storing and managing multiple passwords when protecting Excel workbooks with Aspose.Cells.

using System;
using Aspose.Cells;

// This C# example shows how to create an Excel workbook, assign an opening (encryption) password, protect the workbook structure with a separate password, save the file, attempt to open it without a password, open it with the correct password using LoadOptions, check encryption and structure‑protection flags, handle wrong unprotect passwords, and finally remove the structure protection.
class WorkbookProtectionDemo
{
    static void Main()
    {
        // ---------- Create and protect ----------
        // Create a new workbook and add some data
        Workbook wb = new Workbook();
        wb.Worksheets[0].Cells["A1"].PutValue("Protected Data");

        // Apply an opening (encryption) password
        wb.Settings.Password = "open123";

        // Apply workbook structure protection with its own password
        wb.Protect(ProtectionType.Structure, "struct123");

        // Save the protected workbook
        string filePath = "ProtectedWorkbook.xlsx";
        wb.Save(filePath, SaveFormat.Xlsx);
        wb.Dispose();

        // ---------- Test opening password ----------
        // Attempt to open without providing the opening password (should fail)
        try
        {
            Workbook wbNoPwd = new Workbook(filePath);
            Console.WriteLine("Opened without password (unexpected).");
            wbNoPwd.Dispose();
        }
        catch (Exception ex)
        {
            Console.WriteLine("Failed to open without password: " + ex.Message);
        }

        // Open with the correct opening password
        LoadOptions loadOpts = new LoadOptions { Password = "open123" };
        Workbook loadedWb = new Workbook(filePath, loadOpts);

        // Verify that the workbook was opened successfully
        Console.WriteLine("Workbook opened with password. IsEncrypted: " + loadedWb.Settings.IsEncrypted);

        // Verify that the workbook structure is protected with a password
        Console.WriteLine("Workbook structure protected with password: " + loadedWb.IsWorkbookProtectedWithPassword);

        // ---------- Test structure protection ----------
        // Attempt to unprotect with an incorrect password (should fail)
        try
        {
            loadedWb.Unprotect("wrongPassword");
            Console.WriteLine("Unprotected with wrong password (unexpected).");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Failed to unprotect with wrong password: " + ex.Message);
        }

        // Unprotect the workbook structure with the correct password
        loadedWb.Unprotect("struct123");
        Console.WriteLine("After correct unprotect, IsWorkbookProtectedWithPassword: " + loadedWb.IsWorkbookProtectedWithPassword);

        // Save the workbook after removing structure protection
        string unprotectedPath = "UnprotectedWorkbook.xlsx";
        loadedWb.Save(unprotectedPath, SaveFormat.Xlsx);
        loadedWb.Dispose();
    }
}
