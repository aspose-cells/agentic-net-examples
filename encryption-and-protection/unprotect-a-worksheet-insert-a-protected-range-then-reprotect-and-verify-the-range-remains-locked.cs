// Title: Unprotect Worksheet, Add Password‑Protected Range, Re‑Protect and Verify with Aspose.Cells (.NET)
// Description: This C# example creates a workbook, removes existing protection, defines an AllowEditRanges entry for cells A1:B2 with a password, applies full worksheet protection, saves the file, reloads it, and checks that the range remains password‑locked and the sheet stays protected.
// Keywords: Aspose.Cells | .NET | C# | worksheet unprotect | protected range | AllowEditRanges | password protection | Excel security | save and reload workbook | verify protection
// Common Searches: Aspose.Cells add password protected range C# | unprotect worksheet then protect with password Aspose.Cells | verify protected range after saving workbook | C# example AllowEditRanges Aspose.Cells | how to lock specific cells in Excel using Aspose.Cells
// Developer Intent: Create a password‑secured edit range, protect the sheet, and confirm the settings persist after serialization.
// Use Cases: Distribute a template where only designated cells can be edited by end users | Lock calculation formulas while allowing data entry in input cells | Maintain cell‑level security when a workbook is saved and reopened
// AI Prompts: Generate C# code with Aspose.Cells that unprotects a worksheet, adds an AllowEditRanges entry protected by a password, applies full sheet protection, saves the workbook, reloads it, and prints the range name and protection flags. | Show how to test whether a protected range retains its password and whether the worksheet is still protected after loading a saved file using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsProtectedRangeDemo
{
    // This C# example creates a workbook, removes existing protection, defines an AllowEditRanges entry for cells A1:B2 with a password, applies full worksheet protection, saves the file, reloads it, and checks that the range remains password‑locked and the sheet stays protected.
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet is unprotected before adding a protected range
            worksheet.Unprotect();

            // Add a protected range (cells A1:B2) and set a password for it
            int rangeIndex = worksheet.AllowEditRanges.Add("MyProtectedRange", 0, 0, 1, 1);
            ProtectedRange protectedRange = worksheet.AllowEditRanges[rangeIndex];
            protectedRange.Password = "rangePwd";

            // Protect the worksheet (all protection types) with a worksheet password
            worksheet.Protect(ProtectionType.All, "sheetPwd", null);

            // Save the workbook to a file
            string filePath = "ProtectedRangeDemo.xlsx";
            workbook.Save(filePath);

            // ---------- Load the workbook to verify ----------
            Workbook verifyWorkbook = new Workbook(filePath);
            Worksheet verifySheet = verifyWorkbook.Worksheets[0];

            // Retrieve the previously added protected range
            ProtectedRange verifyRange = verifySheet.AllowEditRanges[0];

            // Output verification results
            Console.WriteLine($"Range Name: {verifyRange.Name}");
            Console.WriteLine($"Is range password protected: {verifyRange.IsProtectedWithPassword}");
            Console.WriteLine($"Worksheet IsProtected: {verifySheet.IsProtected}");
        }
    }
}
