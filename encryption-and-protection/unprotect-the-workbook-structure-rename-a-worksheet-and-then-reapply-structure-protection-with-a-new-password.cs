// Title: Aspose.Cells for .NET – Unprotect Workbook Structure, Rename Worksheet, and Re‑protect with a New Password
// Description: Shows how to create a workbook, apply structure protection, remove it with the original password, rename the first sheet, and then protect the structure again using a different password before saving the file.
// Keywords: Aspose.Cells | C# | .NET | workbook structure protection | unprotect workbook | rename worksheet | change protection password | Workbook.Protect | Workbook.Unprotect | ProtectionType.Structure | Aspose.Cells example
// Common Searches: Aspose.Cells change workbook structure password | rename sheet after unprotecting workbook in C# | re‑apply structure protection with new password Aspose.Cells | how to unprotect and protect workbook structure programmatically | C# example for workbook structure protection rotation
// Developer Intent: Remove existing structure protection, rename a worksheet, and protect the workbook structure again using a new password.
// Use Cases: Update sheet names in a protected workbook without exposing the old password. | Automate periodic password rotation for workbook structure protection. | Prepare a template for distribution by unlocking, renaming sheets, and relocking the structure.
// AI Prompts: Generate a C# snippet that unprotects a workbook's structure, renames multiple worksheets, and then re‑protects it with a new password using Aspose.Cells. | Explain how to change the structure protection password while keeping other protections intact in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, apply structure protection, remove it with the original password, rename the first sheet, and then protect the structure again using a different password before saving the file.
    public class WorkbookStructureProtectionDemo
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Protect the workbook structure with an initial password
            workbook.Protect(ProtectionType.Structure, "oldPassword");

            // Unprotect the workbook structure using the same password
            workbook.Unprotect("oldPassword");

            // Rename the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "RenamedSheet";

            // Re‑apply structure protection with a new password
            workbook.Protect(ProtectionType.Structure, "newPassword");

            // Save the workbook (lifecycle: save)
            workbook.Save("WorkbookStructureProtectionDemo.xlsx");
        }
    }
}
