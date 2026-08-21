// Title: Protect Excel Workbook Structure with a Password using Aspose.Cells for .NET (C#)
// Description: Shows how to lock an entire workbook’s structure—blocking sheet addition, deletion, renaming, moving, or viewing hidden sheets—by applying a strong password with Aspose.Cells for .NET and saving the result.
// Keywords: Aspose.Cells | C# workbook protection | Excel structure password | Protect workbook structure .NET | Workbook.Protect | Excel file security | disable sheet editing | Aspose.Cells encryption | strong password protection | save protected workbook
// Common Searches: Aspose.Cells protect workbook structure C# | How to set password for Excel workbook using Aspose.Cells | Disable sheet addition and deletion in Excel with Aspose | C# code to lock workbook structure Aspose.Cells | Set strong password for Excel file Aspose .NET
// Developer Intent: Apply a password to the workbook’s structure so users cannot add, delete, rename, move, or view hidden worksheets.
// Use Cases: Distribute a template that must remain unchanged by end‑users. | Secure financial or HR reports before emailing to clients. | Prevent accidental sheet modifications in automated Excel generation pipelines.
// AI Prompts: Generate C# code that protects an Aspose.Cells workbook’s structure with a password and also protects the workbook windows. | Provide an example that checks whether a workbook is already structure‑protected before applying a new password. | Explain how to change or remove the workbook structure password using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Shows how to lock an entire workbook’s structure—blocking sheet addition, deletion, renaming, moving, or viewing hidden sheets—by applying a strong password with Aspose.Cells for .NET and saving the result.
class ProtectWorkbookDemo
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // (Optional) Add sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample Data");

        // Protect the entire workbook structure with a strong password
        // This disables adding, deleting, renaming, moving, or viewing hidden sheets
        workbook.Protect(ProtectionType.Structure, "Str0ngP@ssw0rd!");

        // Save the protected workbook to a file
        workbook.Save("ProtectedWorkbook.xlsx", SaveFormat.Xlsx);

        // Release resources
        workbook.Dispose();
    }
}
