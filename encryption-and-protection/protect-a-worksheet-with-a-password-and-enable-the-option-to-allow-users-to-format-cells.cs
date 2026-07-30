// Title: C# – Protect an Excel worksheet with a password while allowing cell formatting using Aspose.Cells
// Description: Creates a workbook, enables the AllowFormattingCell flag on the first worksheet, protects the sheet with a password (ProtectionType.All), and saves the file as ProtectedWorksheet.xlsx.
// Keywords: Aspose.Cells worksheet protection C# | Excel password protect sheet | AllowFormattingCell Aspose | ProtectionType.All example | .NET Excel security | cell formatting on protected sheet
// Common Searches: Aspose.Cells protect sheet with password but allow formatting | C# enable AllowFormattingCell while protecting worksheet | How to set worksheet protection options in Aspose.Cells .NET | Excel sheet password protection allowing cell style changes
// Developer Intent: Secure a worksheet with a password yet permit users to format cells.
// Use Cases: Distribute a read‑only report where users can change fonts or colors without editing data. | Provide a template that blocks data entry but allows styling adjustments such as bold or background fills. | Share a workbook where protection prevents content changes but formatting actions like borders remain available.
// AI Prompts: Generate C# code that protects an Aspose.Cells worksheet with a password and enables only cell formatting. | Show how to set AllowFormattingCell = true and apply ProtectionType.All on a worksheet using Aspose.Cells. | Explain how to modify worksheet protection to allow additional actions (e.g., inserting rows) while keeping the sheet password protected.

using System;
using Aspose.Cells;

// Creates a workbook, enables the AllowFormattingCell flag on the first worksheet, protects the sheet with a password (ProtectionType.All), and saves the file as ProtectedWorksheet.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Get the protection settings for the worksheet
        Protection protection = sheet.Protection;

        // Enable formatting of cells while the sheet is protected
        protection.AllowFormattingCell = true;

        // Protect the worksheet with a password and allow all protection types
        sheet.Protect(ProtectionType.All, "myPassword", null);

        // Save the workbook
        workbook.Save("ProtectedWorksheet.xlsx");
    }
}
