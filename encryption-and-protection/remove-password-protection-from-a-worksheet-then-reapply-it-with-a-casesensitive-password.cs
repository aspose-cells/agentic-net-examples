// Title: Remove and Re‑apply Worksheet Protection with a Case‑Sensitive Password using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, protect its first worksheet with an initial password, unprotect it, and then protect the same worksheet again using a case‑sensitive password. The final file is saved as ProtectedWorksheet.xlsx.
// Keywords: Aspose.Cells | C# | worksheet protection | remove worksheet password | case‑sensitive password | unprotect worksheet | protect worksheet | ProtectionType.All | Excel encryption | Aspose.Cells API
// Common Searches: Aspose.Cells remove worksheet password C# | how to unprotect a worksheet and set a new password Aspose.Cells | case‑sensitive worksheet protection Aspose.Cells .NET | protect Excel sheet with case‑sensitive password using Aspose | re‑apply worksheet protection after unprotecting Aspose.Cells
// Developer Intent: Remove existing protection from a worksheet and then apply a new, case‑sensitive password using the Aspose.Cells library.
// Use Cases: Reset a temporary password after a user changes credentials in an automated reporting pipeline. | Apply stronger, case‑sensitive protection before distributing a workbook to external partners. | Programmatically rotate worksheet passwords for compliance with security policies.
// AI Prompts: Write C# code with Aspose.Cells that unprotects a worksheet using a known password and then protects it again with a case‑sensitive password. | Explain whether Aspose.Cells treats worksheet passwords as case‑sensitive and show a short example. | Provide error‑handling code for Worksheet.Unprotect when an incorrect password is supplied in Aspose.Cells.

using Aspose.Cells;
using System;

// Demonstrates how to create a workbook, protect its first worksheet with an initial password, unprotect it, and then protect the same worksheet again using a case‑sensitive password. The final file is saved as ProtectedWorksheet.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Protect the worksheet with an initial password
        string initialPassword = "Password123";
        worksheet.Protect(ProtectionType.All, initialPassword, null);

        // Remove the protection using the same password
        worksheet.Unprotect(initialPassword);

        // Re‑apply protection with a case‑sensitive password
        string caseSensitivePassword = "MyCaseSensitivePass";
        worksheet.Protect(ProtectionType.All, caseSensitivePassword, null);

        // Save the workbook
        workbook.Save("ProtectedWorksheet.xlsx");
    }
}
