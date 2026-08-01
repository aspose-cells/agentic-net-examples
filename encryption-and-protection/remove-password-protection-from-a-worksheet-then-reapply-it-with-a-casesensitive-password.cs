// Title: Remove and Re‑apply Worksheet Protection with a Case‑Sensitive Password – Aspose.Cells for .NET
// Description: Demonstrates how to protect a worksheet, unprotect it with the original password, and then protect it again using a mixed‑case password (case‑sensitive by default) with Aspose.Cells for .NET, and finally save the workbook.
// Keywords: Aspose.Cells worksheet unprotect | C# protect worksheet case sensitive | remove worksheet password Aspose.Cells | re‑apply worksheet protection .NET | Aspose.Cells password case sensitivity
// Common Searches: Aspose.Cells unprotect worksheet C# | how to change worksheet password Aspose.Cells | case sensitive worksheet protection .NET | remove and reset worksheet password programmatically | Aspose.Cells protect worksheet with mixed case password
// Developer Intent: Programmatically remove existing worksheet protection and then apply a new case‑sensitive password using Aspose.Cells for .NET.
// Use Cases: Update an outdated worksheet password to meet new security policies. | Apply a stronger, mixed‑case password before distributing a workbook to external partners. | Temporarily lift protection for data manipulation and automatically restore secure protection afterward.
// AI Prompts: Write C# code with Aspose.Cells that unprotects a worksheet using a known password and then protects it again with a new mixed‑case password. | Explain Aspose.Cells' handling of password case sensitivity for worksheet protection and show how to change the password via code. | Provide a step‑by‑step tutorial for removing worksheet protection, confirming the unprotected state, and re‑applying protection with a case‑sensitive password in Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to protect a worksheet, unprotect it with the original password, and then protect it again using a mixed‑case password (case‑sensitive by default) with Aspose.Cells for .NET, and finally save the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Protect the worksheet with an initial password (case‑insensitive example)
        worksheet.Protect(ProtectionType.All, "Password123", null);
        Console.WriteLine("Worksheet initially protected: " + worksheet.IsProtected);

        // Remove the existing protection using the correct password
        worksheet.Unprotect("Password123");
        Console.WriteLine("Worksheet after unprotect: " + worksheet.IsProtected);

        // Re‑apply protection with a case‑sensitive password.
        // Passwords in Aspose.Cells are case‑sensitive by default, so using mixed case ensures it.
        worksheet.Protect(ProtectionType.All, "MySecretPASS", null);
        Console.WriteLine("Worksheet re‑protected with case‑sensitive password: " + worksheet.IsProtected);

        // Save the workbook to a file
        workbook.Save("ProtectedWorksheet.xlsx");
    }
}
