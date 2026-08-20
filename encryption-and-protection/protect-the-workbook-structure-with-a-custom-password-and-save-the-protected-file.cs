// Title: Protect Workbook Structure with a Password using Aspose.Cells for .NET
// Description: Shows how to instantiate a Workbook, apply structure‑only protection with a custom password via Workbook.Protect, and save the protected file as an .xlsx document.
// Keywords: Aspose.Cells | .NET | C# workbook protection | structure password Excel | Workbook.Protect | Excel file security | protect workbook without sheet lock | password‑protected workbook
// Common Searches: Aspose.Cells protect workbook structure C# | How to add a structure password to an Excel file using .NET | Example of Workbook.Protect with custom password | Save a password‑protected workbook with Aspose.Cells | Protect only workbook layout, not sheets, in C#
// Developer Intent: Apply a password that locks the workbook’s structure and save the resulting file.
// Use Cases: Create a template that prevents users from adding, deleting, or renaming sheets. | Distribute a report where the sheet order must remain unchanged. | Share a workbook externally while keeping its layout immutable.
// AI Prompts: Generate C# code that protects only the workbook structure with a password using Aspose.Cells. | Extend the example to also protect workbook windows in addition to the structure. | Describe how to programmatically verify that structure protection is enabled after saving.

using System;
using Aspose.Cells;

// Shows how to instantiate a Workbook, apply structure‑only protection with a custom password via Workbook.Protect, and save the protected file as an .xlsx document.
class ProtectWorkbookStructure
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Protect the workbook's structure with a custom password
        workbook.Protect(ProtectionType.Structure, "MySecretPassword");

        // Save the protected workbook to a file
        workbook.Save("ProtectedWorkbook.xlsx", SaveFormat.Xlsx);

        // Release resources
        workbook.Dispose();
    }
}
