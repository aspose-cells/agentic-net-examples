// Title: Protect Workbook Structure with a Custom Password using Aspose.Cells for .NET (C#)
// Description: Creates a new Workbook, applies structure protection with a user‑defined password via wb.Protect(ProtectionType.Structure, "MySecretPassword"), saves the file as a password‑protected .xlsx, and releases resources.
// Keywords: Aspose.Cells C# protect workbook structure | Workbook.Protect method example | custom password Excel protection .NET | save password‑protected XLSX with Aspose.Cells | Excel file encryption Aspose.Cells
// Common Searches: Aspose.Cells protect workbook structure C# | set custom password for Excel workbook using Aspose | how to save a password‑protected .xlsx with Aspose.Cells | C# code to lock workbook structure Aspose.Cells | protect Excel file without protecting worksheets
// Developer Intent: Apply a custom password to lock the workbook's structure and write the protected file to disk.
// Use Cases: Generate a template workbook that prevents users from adding, moving, or deleting sheets. | Distribute financial reports where the sheet order must remain unchanged. | Automate creation of corporate Excel files that enforce a fixed layout across all recipients.
// AI Prompts: Write C# code with Aspose.Cells to protect only the workbook structure using a specified password and save as .xlsx. | Show how to protect a workbook's structure and also set an opening password with Aspose.Cells for .NET. | Explain the difference between ProtectionType.Structure and ProtectionType.Windows in Aspose.Cells and provide code samples for each.

using System;
using Aspose.Cells;

// Creates a new Workbook, applies structure protection with a user‑defined password via wb.Protect(ProtectionType.Structure, "MySecretPassword"), saves the file as a password‑protected .xlsx, and releases resources.
class ProtectWorkbook
{
    static void Main()
    {
        // Create a new workbook (empty by default)
        Workbook wb = new Workbook();

        // Protect the workbook's structure with a custom password
        wb.Protect(ProtectionType.Structure, "MySecretPassword");

        // Save the protected workbook to a file
        wb.Save("ProtectedWorkbook.xlsx", SaveFormat.Xlsx);

        // Release resources
        wb.Dispose();
    }
}
