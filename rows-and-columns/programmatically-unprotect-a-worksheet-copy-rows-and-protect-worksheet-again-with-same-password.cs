// Title: C# – Unprotect Worksheet, Copy Rows, and Re‑protect with Same Password using Aspose.Cells
// Description: Shows how to protect a worksheet with a password, temporarily remove the protection, copy rows while keeping data and formatting, and then re‑apply the original password in Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | worksheet protection | unprotect worksheet | copy rows | CopyRows method | ProtectionType.All | password protection | preserve formatting | Excel automation
// Common Searches: Aspose.Cells unprotect worksheet copy rows | How to copy rows in a protected sheet using Aspose.Cells | Re‑apply worksheet password after copying rows Aspose.Cells .NET | CopyRows example Aspose.Cells | Temporarily remove protection to modify Excel with Aspose.Cells
// Developer Intent: Remove worksheet protection, duplicate selected rows, and restore the same password afterward.
// Use Cases: Create a locked template, then programmatically duplicate header rows for new sections without exposing the password. | Automate monthly reports where a protected sheet needs row duplication while maintaining original security settings. | Migrate data inside a secured workbook by unprotecting, moving rows, and re‑securing the sheet in a single routine.
// AI Prompts: Generate C# code with Aspose.Cells that removes worksheet protection, copies a range of rows preserving formatting, and re‑applies the original password. | Provide an Aspose.Cells .NET example demonstrating the CopyRows method on a sheet that is temporarily unprotected. | Show how to protect a worksheet with ProtectionType.All, unprotect it, duplicate rows, and protect it again using the same password in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsWorksheetProtectionDemo
{
    // Shows how to protect a worksheet with a password, temporarily remove the protection, copy rows while keeping data and formatting, and then re‑apply the original password in Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data in rows 0 and 1
            sheet.Cells["A1"].PutValue("Row 1");
            sheet.Cells["A2"].PutValue("Row 2");

            // Define the protection password
            string password = "MySecretPwd";

            // Protect the worksheet with the password (all protection types)
            sheet.Protect(ProtectionType.All, password, null);

            // Unprotect the worksheet using the same password
            sheet.Unprotect(password);

            // Copy rows 0-1 to rows 2-3 (preserving data and formatting)
            // Parameters: source cells, source start row, destination start row, number of rows to copy
            sheet.Cells.CopyRows(sheet.Cells, 0, 2, 2);

            // Re‑protect the worksheet with the original password
            sheet.Protect(ProtectionType.All, password, null);

            // Save the workbook to a file
            workbook.Save("ProtectedAndCopiedRows.xlsx");
        }
    }
}
