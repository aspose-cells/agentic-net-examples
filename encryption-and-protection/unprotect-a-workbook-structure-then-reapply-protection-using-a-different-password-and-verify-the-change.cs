// Title: Change workbook structure password with Aspose.Cells for .NET (C#) – unprotect, re‑protect, verify
// Description: Shows how to protect an Excel workbook's structure with a password, remove that protection, apply a new password, and confirm the new password works using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# workbook protection | unprotect workbook structure | change Excel password | protect workbook structure | verify workbook protection | Aspose.Cells Protect | Aspose.Cells Unprotect | Excel password rotation | .NET Excel security example
// Common Searches: Aspose.Cells change workbook password C# | unprotect workbook structure Aspose.Cells | re‑protect Excel file with new password using Aspose | verify workbook protection after password change | C# code to modify workbook structure protection
// Developer Intent: Remove the current structure protection, set a new password, and validate that the new password can unlock the workbook.
// Use Cases: Rotate workbook structure passwords to comply with security policies. | Replace an old password before sharing the file with a new user group. | Automate a validation step that confirms a workbook remains protected after a password update.
// AI Prompts: Generate C# code with Aspose.Cells that unprotects a workbook structure, applies a new password, and verifies the change. | Explain step‑by‑step how to remove and re‑apply structure protection on an Excel file using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Shows how to protect an Excel workbook's structure with a password, remove that protection, apply a new password, and confirm the new password works using Aspose.Cells for .NET.
class WorkbookProtectionDemo
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook wb = new Workbook();
        Worksheet sheet = wb.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample Data");

        // Protect the workbook structure with an initial password
        string oldPassword = "oldPass123";
        wb.Protect(ProtectionType.Structure, oldPassword);
        Console.WriteLine("Initially protected: " + wb.IsWorkbookProtectedWithPassword); // Expected: True

        // Save the protected workbook
        string protectedPath = "protected_workbook.xlsx";
        wb.Save(protectedPath);
        wb.Dispose();

        // Load the workbook (structure protection does not require a load password)
        Workbook loadedWb = new Workbook(protectedPath);
        Console.WriteLine("After load, protected: " + loadedWb.IsWorkbookProtectedWithPassword); // Expected: True

        // Unprotect the workbook using the original password
        loadedWb.Unprotect(oldPassword);
        Console.WriteLine("After unprotect, protected: " + loadedWb.IsWorkbookProtectedWithPassword); // Expected: False

        // Re‑apply protection with a new password
        string newPassword = "newPass456";
        loadedWb.Protect(ProtectionType.Structure, newPassword);
        Console.WriteLine("After re‑protect, protected: " + loadedWb.IsWorkbookProtectedWithPassword); // Expected: True

        // Verify that the new password works by attempting to unprotect with it
        loadedWb.Unprotect(newPassword);
        Console.WriteLine("After verifying new password, protected: " + loadedWb.IsWorkbookProtectedWithPassword); // Expected: False

        // Save the final workbook
        string finalPath = "reprotected_workbook.xlsx";
        loadedWb.Save(finalPath);
        loadedWb.Dispose();
    }
}
