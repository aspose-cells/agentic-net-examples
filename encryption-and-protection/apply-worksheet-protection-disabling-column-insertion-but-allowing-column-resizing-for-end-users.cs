// Title: Aspose.Cells C# – Protect Worksheet: Block Column Insertion, Allow Column Resizing
// Description: Demonstrates how to create a workbook, enable worksheet protection, disable column insertion (AllowInsertingColumn = false), permit column resizing (AllowFormattingColumn = true), set a password, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells worksheet protection C# | disable column insertion | allow column resizing | worksheet password protection | ProtectionType.All Aspose.Cells
// Common Searches: Aspose.Cells prevent column insert but allow resize | C# protect worksheet with specific permissions Aspose.Cells | how to set worksheet password and allow formatting in Aspose.Cells
// Developer Intent: Protect a worksheet so users cannot insert new columns while still being able to change column widths.
// Use Cases: Template distribution where layout must stay fixed but users can adjust column widths for readability. | Financial report sharing that requires column structure integrity with flexible printing layout. | Collaborative workbook secured with a password, granting formatting rights but blocking structural changes.
// AI Prompts: Write C# code with Aspose.Cells to protect a worksheet, disable column insertion, enable column resizing, and apply a password. | Explain how to extend the protection settings to also block row deletion while keeping column formatting enabled. | Show how to apply the same protection configuration to all worksheets in a workbook using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, enable worksheet protection, disable column insertion (AllowInsertingColumn = false), permit column resizing (AllowFormattingColumn = true), set a password, and save the file using Aspose.Cells for .NET.
class WorksheetProtectionDemo
{
    public static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Access the worksheet's protection settings
        Protection protection = sheet.Protection;

        // Disable column insertion while the sheet is protected
        protection.AllowInsertingColumn = false;

        // Enable column formatting (allows users to resize columns)
        protection.AllowFormattingColumn = true;

        // Optional: set a password for the protection
        protection.Password = "pwd123";

        // Apply protection with all protection types enabled
        sheet.Protect(ProtectionType.All);

        // Save the workbook
        workbook.Save("WorksheetProtection.xlsx");
    }
}
