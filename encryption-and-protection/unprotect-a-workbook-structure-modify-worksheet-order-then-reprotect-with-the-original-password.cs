// Title: Unprotect, Reorder Worksheets, and Re‑protect Excel Workbook Structure with Aspose.Cells for .NET
// Description: Load a password‑protected Excel file, call Workbook.Unprotect to lift structure protection, move worksheets (e.g., shift the first sheet to the end) using Worksheet.MoveTo, then re‑apply Workbook.Protect with the original password and save the result.
// Keywords: Aspose.Cells C# unprotect workbook | Excel workbook structure protection | Workbook.Unprotect example | Workbook.Protect structure password | move worksheet Aspose.Cells | reorder sheets programmatically | C# Excel automation | protect workbook structure .NET | Aspose.Cells worksheet ordering | Excel file password handling
// Common Searches: how to unprotect workbook structure using Aspose.Cells C# | reorder worksheets in a protected Excel file Aspose.Cells | move first sheet to last position C# Aspose.Cells | protect Excel workbook structure with password .NET | Aspose.Cells example for changing sheet order while keeping protection
// Developer Intent: Remove structure protection, change the order of worksheets, and restore the same protection password.
// Use Cases: Adjust sheet sequence in a template that must stay locked before distribution. | Batch‑process multiple workbooks to standardize tab order while preserving structure security. | Automate re‑ordering of data sheets after import without exposing the workbook to accidental edits.
// AI Prompts: Generate C# code with Aspose.Cells to unprotect a workbook's structure, move a specific worksheet to a new index, and re‑protect it using the same password. | Explain error handling for an incorrect password when calling Workbook.Unprotect in Aspose.Cells. | Show how to reorder several worksheets in a protected workbook while keeping all existing protection settings intact.

using System;
using Aspose.Cells;

// Load a password‑protected Excel file, call Workbook.Unprotect to lift structure protection, move worksheets (e.g., shift the first sheet to the end) using Worksheet.MoveTo, then re‑apply Workbook.Protect with the original password and save the result.
class WorkbookStructureReorder
{
    static void Main()
    {
        // Path to the protected workbook and the password used for protection
        string inputFile = "input.xlsx";
        string outputFile = "output.xlsx";
        string password = "myPassword";

        // Load the existing workbook (create/load rule)
        Workbook workbook = new Workbook(inputFile);

        // Unprotect the workbook structure using the original password
        workbook.Unprotect(password);

        // Example of reordering worksheets:
        // Move the first worksheet to the last position
        Worksheet firstSheet = workbook.Worksheets[0];
        firstSheet.MoveTo(workbook.Worksheets.Count - 1);

        // Re‑protect the workbook structure with the same password
        workbook.Protect(ProtectionType.Structure, password);

        // Save the modified workbook (save rule)
        workbook.Save(outputFile);
    }
}
