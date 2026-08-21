// Title: C# – Unprotect Workbook Structure, Rename Worksheet, and Re‑protect with New Password using Aspose.Cells
// Description: Shows how to load an Excel file with Aspose.Cells for .NET, remove its existing structure protection, rename a worksheet, and then apply structure protection again using a different password before saving.
// Keywords: Aspose.Cells | C# | unprotect workbook structure | rename worksheet | protect workbook structure | new password | Excel protection programmatically | Workbook.Unprotect | Workbook.Protect | ProtectionType.Structure
// Common Searches: Aspose.Cells rename sheet after unprotect | Change workbook structure password C# | How to unprotect and protect Excel workbook with Aspose.Cells | C# code to rename worksheet in protected workbook | Update sheet name in password‑protected Excel using Aspose.Cells
// Developer Intent: Remove the current structure protection, change a sheet’s name, and re‑apply structure protection with a different password in an Excel workbook via C#.
// Use Cases: Modify sheet names in a template that is distributed with structure protection. | Rotate workbook structure passwords automatically to meet security policies. | Rename default sheets in generated reports before delivering them to end users.
// AI Prompts: Write C# code using Aspose.Cells to unprotect a workbook’s structure, rename a specific worksheet, and protect the structure again with a new password. | Explain how Workbook.Protect with ProtectionType.Structure works and how to change the protection password after calling Workbook.Unprotect. | Provide robust error‑handling for an incorrect old password when invoking Workbook.Unprotect in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    // Shows how to load an Excel file with Aspose.Cells for .NET, remove its existing structure protection, rename a worksheet, and then apply structure protection again using a different password before saving.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Unprotect the workbook structure using the current password
            workbook.Unprotect("oldPassword");

            // Rename the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "RenamedSheet";

            // Re‑apply structure protection with a new password
            workbook.Protect(ProtectionType.Structure, "newPassword");

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}
