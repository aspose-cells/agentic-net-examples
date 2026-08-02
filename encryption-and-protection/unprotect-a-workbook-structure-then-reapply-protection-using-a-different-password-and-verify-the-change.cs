// Title: Change Workbook Structure Protection Password with Aspose.Cells (C#)
// Description: Demonstrates how to protect a workbook's structure with an initial password, remove the protection, re‑apply it using a new password, verify the protection flags, and save both versions using Aspose.Cells for .NET.
// Keywords: Aspose.Cells protect workbook structure | Aspose.Cells unprotect workbook | change workbook password C# | Workbook.IsWorkbookProtectedWithPassword | Workbook.Settings.IsProtected | C# Excel protection password rotation | Aspose.Cells Protect method | Aspose.Cells Unprotect method
// Common Searches: How to change the password of a protected Excel workbook with Aspose.Cells | Unprotect workbook structure using Aspose.Cells C# | Verify workbook protection status after Unprotect in Aspose.Cells | Replace old workbook password with new one in .NET | Aspose.Cells protect and unprotect example
// Developer Intent: The developer needs to remove existing structure protection from an Excel workbook, apply a new password, and confirm the protection state before and after the change.
// Use Cases: Rotate workbook passwords for security compliance before distributing the file. | Programmatically modify sheet order after unprotecting, then re‑secure the workbook. | Create automated tests that validate protection flags when changing passwords.
// AI Prompts: Generate C# code with Aspose.Cells that unprotects a workbook structure using an old password, then protects it again with a new password and prints the protection status. | Explain the difference between Workbook.IsWorkbookProtectedWithPassword and Workbook.Settings.IsProtected after calling Unprotect and Protect. | Write a C# unit test using NUnit that asserts the workbook is unprotected after Unprotect and protected after Protect with a new password.

using System;
using Aspose.Cells;

namespace AsposeCellsWorkbookProtectionDemo
{
    // Demonstrates how to protect a workbook's structure with an initial password, remove the protection, re‑apply it using a new password, verify the protection flags, and save both versions using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Protect the workbook structure with an initial password
            string oldPassword = "oldPass";
            workbook.Protect(ProtectionType.Structure, oldPassword);
            Console.WriteLine("Workbook protected with old password: " + workbook.IsWorkbookProtectedWithPassword);

            // Save the initially protected workbook
            workbook.Save("protected_workbook.xlsx");

            // Unprotect the workbook using the old password
            workbook.Unprotect(oldPassword);
            Console.WriteLine("Workbook unprotected. Is protected with password? " + workbook.IsWorkbookProtectedWithPassword);
            Console.WriteLine("Workbook settings IsProtected? " + workbook.Settings.IsProtected);

            // Re‑apply protection with a new password
            string newPassword = "newPass";
            workbook.Protect(ProtectionType.Structure, newPassword);
            Console.WriteLine("Workbook re‑protected with new password: " + workbook.IsWorkbookProtectedWithPassword);
            Console.WriteLine("Workbook settings IsProtected after re‑protect? " + workbook.Settings.IsProtected);

            // Save the re‑protected workbook
            workbook.Save("reprotected_workbook.xlsx");
        }
    }
}
