// Title: C# – Unprotect and Re‑protect Workbook Structure with a New Password using Aspose.Cells for .NET
// Description: Demonstrates how to remove structure protection from an Aspose.Cells Workbook, apply a new password, verify that the old password no longer works, confirm the new password unlocks the file, and save the workbook before disposal.
// Keywords: Aspose.Cells protect workbook structure | unprotect workbook C# | change workbook password Aspose.Cells | verify workbook protection .NET | Excel workbook structure password | C# Aspose.Cells example
// Common Searches: change password of protected Excel workbook using Aspose.Cells | unprotect and re‑protect workbook structure C# | verify old password fails after re‑protecting workbook | Aspose.Cells workbook protection sample
// Developer Intent: Remove existing structure protection, set a new password, and ensure the previous password is invalid while the new one works.
// Use Cases: Automate password rotation for Excel workbooks in CI/CD pipelines to comply with security policies. | Update workbook protection before sharing files with external partners after a review cycle. | Batch process reports, assigning a unique password to each workbook and confirming old passwords are rejected.
// AI Prompts: Write C# code with Aspose.Cells that unprotects a workbook structure, applies a new password, and includes error handling for an incorrect old password. | Explain how to test that a workbook no longer accepts its previous password after re‑protecting it with Aspose.Cells. | Provide a short example showing how to save a workbook before and after changing its protection password, and how to check the protection state.

using System;
using Aspose.Cells;

namespace AsposeCellsWorkbookProtectionDemo
{
    // Demonstrates how to remove structure protection from an Aspose.Cells Workbook, apply a new password, verify that the old password no longer works, confirm the new password unlocks the file, and save the workbook before disposal.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Protect the workbook structure with an initial password
            string oldPassword = "oldPass123";
            workbook.Protect(ProtectionType.Structure, oldPassword);
            Console.WriteLine("Workbook initially protected: " + workbook.IsWorkbookProtectedWithPassword);

            // Save the initially protected workbook (optional, just for demonstration)
            workbook.Save("InitiallyProtected.xlsx");

            // Unprotect the workbook using the old password
            workbook.Unprotect(oldPassword);
            Console.WriteLine("Workbook unprotected: " + !workbook.IsWorkbookProtectedWithPassword);

            // Re‑apply protection with a new password
            string newPassword = "newPass456";
            workbook.Protect(ProtectionType.Structure, newPassword);
            Console.WriteLine("Workbook re‑protected with new password: " + workbook.IsWorkbookProtectedWithPassword);

            // Verify that the new password works and the old one does not
            try
            {
                // Attempt to unprotect with the old password (should fail)
                workbook.Unprotect(oldPassword);
                Console.WriteLine("Unexpectedly unprotected with old password.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to unprotect with old password (expected): " + ex.Message);
            }

            // Unprotect with the new password to confirm it works
            workbook.Unprotect(newPassword);
            Console.WriteLine("Successfully unprotected with new password: " + !workbook.IsWorkbookProtectedWithPassword);

            // Save the final workbook
            workbook.Save("FinalProtectedWorkbook.xlsx");

            // Clean up
            workbook.Dispose();
        }
    }
}
