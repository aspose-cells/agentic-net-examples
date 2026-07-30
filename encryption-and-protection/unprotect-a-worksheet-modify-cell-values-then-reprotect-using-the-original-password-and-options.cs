// Title: C# – Unprotect, edit, and re‑protect an Excel worksheet with the original password using Aspose.Cells
// Description: Demonstrates how to apply full protection to a worksheet, store its settings, unprotect it, modify cells, and then re‑apply the same password and protection options before saving the workbook with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# worksheet protection | unprotect worksheet | reprotect worksheet | Excel password protection | modify cells after unprotect | preserve protection settings | ProtectionType.All | Aspose.Cells example
// Common Searches: Aspose.Cells unprotect worksheet C# | How to edit a protected Excel sheet with Aspose.Cells | Re‑apply original protection after modifying cells Aspose.Cells | Store and reuse worksheet protection settings .NET | C# code to protect and unprotect Excel worksheet using password
// Developer Intent: Temporarily remove worksheet protection, update specific cells, and then restore the original password and protection flags programmatically.
// Use Cases: Update totals in a locked financial report, then re‑lock the sheet to prevent further edits. | Batch‑process a template where only certain cells need to be changed while keeping the original protection intact. | Create an automated routine that edits designated cells in a protected worksheet and restores all original protection options afterward.
// AI Prompts: Generate C# code with Aspose.Cells that unprotects a worksheet, edits a range of cells, and re‑protects it preserving the original password and protection flags. | Explain step‑by‑step how to capture and reuse worksheet protection settings when unprotecting and protecting an Excel sheet using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Demonstrates how to apply full protection to a worksheet, store its settings, unprotect it, modify cells, and then re‑apply the same password and protection options before saving the workbook with Aspose.Cells for .NET.
class UnprotectModifyReprotectDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // ----- Set initial protection options and protect the sheet -----
        Protection originalProtection = sheet.Protection;
        originalProtection.AllowEditingContent = false;   // disallow editing locked cells
        originalProtection.AllowEditingObject = false;    // disallow editing objects
        originalProtection.Password = "mySecretPwd";      // set password

        // Apply protection with the password and all protection types
        sheet.Protect(ProtectionType.All, originalProtection.Password, null);

        // ----- Store original protection settings before unprotecting -----
        bool allowEditingContent = originalProtection.AllowEditingContent;
        bool allowEditingObject = originalProtection.AllowEditingObject;
        string password = originalProtection.Password;

        // ----- Unprotect the worksheet using the stored password -----
        sheet.Unprotect(password);

        // ----- Modify cell values while the sheet is unprotected -----
        sheet.Cells["A1"].PutValue("Updated Value");
        sheet.Cells["B2"].PutValue(9876);

        // ----- Re‑apply the original protection options -----
        Protection reprotect = sheet.Protection;
        reprotect.AllowEditingContent = allowEditingContent;
        reprotect.AllowEditingObject = allowEditingObject;
        reprotect.Password = password;

        // Protect the worksheet again with the same password and protection type
        sheet.Protect(ProtectionType.All, password, null);

        // Save the workbook
        workbook.Save("UnprotectModifyReprotect.xlsx");
    }
}
