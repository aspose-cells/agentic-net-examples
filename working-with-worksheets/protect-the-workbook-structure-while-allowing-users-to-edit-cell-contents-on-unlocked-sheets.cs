// Title: Protect Workbook Structure Only with Aspose.Cells for .NET (Sheets Remain Editable)
// Description: Shows how to create a workbook, optionally add a worksheet, apply structure‑only protection with a password via Workbook.Protect(ProtectionType.Structure,…), and save the file. The protection blocks adding, deleting, or renaming worksheets while all cell contents stay editable.
// Keywords: Aspose.Cells | C# workbook structure protection | .NET Excel protection | Protect Structure Aspose.Cells | password‑protected workbook | editable cells Excel template | prevent sheet addition deletion | Excel workbook lock structure
// Common Searches: Aspose.Cells protect only workbook structure C# | prevent adding or deleting sheets with Aspose.Cells .NET | keep cells editable while locking workbook structure | Excel template protect sheet order Aspose.Cells
// Developer Intent: The developer wants to lock the workbook’s structure (add/delete/rename sheets) but still allow users to edit any cell content.
// Use Cases: Distribute a multi‑sheet template where users can enter data but cannot change sheet order or names. | Share a financial model that must retain its worksheet layout while remaining fully editable. | Create a report package that prevents accidental sheet insertion or removal yet allows content updates.
// AI Prompts: Generate C# code that protects only the workbook structure with a password using Aspose.Cells and later unprotects it. | Provide an example that combines workbook structure protection with cell‑level locking/unlocking in Aspose.Cells for .NET. | Explain how to apply structure‑only protection together with worksheet‑level protection in Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to create a workbook, optionally add a worksheet, apply structure‑only protection with a password via Workbook.Protect(ProtectionType.Structure,…), and save the file. The protection blocks adding, deleting, or renaming worksheets while all cell contents stay editable.
class ProtectWorkbookStructure
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add an additional worksheet (optional)
        workbook.Worksheets.Add("SecondSheet");

        // Protect only the workbook structure with a password.
        // This prevents adding, deleting, or renaming worksheets,
        // while leaving the worksheets themselves editable.
        workbook.Protect(ProtectionType.Structure, "MySecretPassword");

        // Save the protected workbook
        workbook.Save("ProtectedWorkbookStructure.xlsx");
    }
}
