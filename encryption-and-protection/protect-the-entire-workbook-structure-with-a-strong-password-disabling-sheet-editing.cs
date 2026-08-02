// Title: Protect Workbook Structure and All Sheets with a Strong Password using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, optionally add data, protect every worksheet with ProtectionType.All, secure the workbook's structure with ProtectionType.Structure, and save the file as an .xlsx using a strong password in Aspose.Cells for .NET.
// Keywords: Aspose.Cells protect workbook structure | protect all worksheets password .NET | Aspose.Cells workbook protection example | set strong password Excel file Aspose | disable sheet editing Aspose.Cells | ProtectionType.All Aspose.Cells | ProtectionType.Structure Aspose.Cells
// Common Searches: Aspose.Cells protect workbook structure with password | C# code to lock all worksheets in Excel using Aspose | How to prevent adding or moving sheets in Aspose.Cells | Secure Excel file with strong password in .NET | Aspose.Cells example for workbook and sheet protection
// Developer Intent: Apply a strong password to lock every worksheet and the workbook's structure, preventing cell edits and sheet modifications.
// Use Cases: Distribute a read‑only financial report that cannot be altered or rearranged. | Provide a template where only specific ranges are editable while the rest of the workbook remains locked. | Protect confidential data before publishing an Excel file to external stakeholders.
// AI Prompts: Generate C# code with Aspose.Cells that protects all worksheets using ProtectionType.All and secures the workbook structure with the same password. | Explain the difference between ProtectionType.All and ProtectionType.Structure in Aspose.Cells for .NET and how they combine to block editing and sheet management.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, optionally add data, protect every worksheet with ProtectionType.All, secure the workbook's structure with ProtectionType.Structure, and save the file as an .xlsx using a strong password in Aspose.Cells for .NET.
class ProtectWorkbookDemo
{
    public static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // (Optional) Add sample data to the first worksheet
        Worksheet firstSheet = workbook.Worksheets[0];
        firstSheet.Cells["A1"].PutValue("Sample Data");

        // Protect every worksheet to prevent editing of cells, formulas, etc.
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Protect with all protection types and a strong password
            sheet.Protect(ProtectionType.All, "StrongPassword!@#123", null);
        }

        // Protect the workbook structure (add/delete/move sheets) with the same strong password
        workbook.Protect(ProtectionType.Structure, "StrongPassword!@#123");

        // Save the workbook; the structure and sheets are now protected
        workbook.Save("ProtectedWorkbook.xlsx", SaveFormat.Xlsx);

        // Release resources
        workbook.Dispose();
    }
}
