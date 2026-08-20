// Title: Aspose.Cells .NET: Unprotect Sheet, Add Password‑Protected Range, Re‑protect and Verify
// Description: Demonstrates how to unprotect a worksheet, create a password‑protected range (A1:B2), protect the sheet with its own password, save the workbook, reload it, and confirm that both the sheet and the range retain their protection using Aspose.Cells for .NET.
// Keywords: Aspose.Cells protect worksheet | password protected range C# | verify protected range after sheet protection | save workbook with protected range | load workbook check protection | Aspose.Cells .NET example
// Common Searches: add password protected range Aspose.Cells .NET | check if protected range stays locked after sheet protection | persist protected ranges when saving workbook Aspose.Cells | unprotect worksheet then protect with range password
// Developer Intent: Add a password‑protected cell range, protect the worksheet, and ensure the protection persists after saving and reloading the file.
// Use Cases: Create templates where only specific cells are editable while the rest of the sheet is locked. | Generate reports that require immutable cells even when the worksheet is globally protected. | Validate that existing protected ranges remain active after a workbook is opened again.
// AI Prompts: Generate C# code with Aspose.Cells to unprotect a worksheet, add a password‑protected range, protect the sheet, save, reload, and verify the protection. | Explain how Aspose.Cells stores protected‑range passwords and how to read IsProtectedWithPassword after loading a workbook. | Suggest robust error‑handling patterns for worksheet and range protection operations in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsProtectedRangeDemo
{
    // Demonstrates how to unprotect a worksheet, create a password‑protected range (A1:B2), protect the sheet with its own password, save the workbook, reload it, and confirm that both the sheet and the range retain their protection using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet is unprotected (in case it was previously protected)
            worksheet.Unprotect();

            // Add a protected range (A1:B2) and set a password for the range
            int rangeIndex = worksheet.AllowEditRanges.Add("MyProtectedRange", 0, 0, 1, 1);
            ProtectedRange protectedRange = worksheet.AllowEditRanges[rangeIndex];
            protectedRange.Password = "rangePassword";

            // Protect the worksheet with its own password
            worksheet.Protect(ProtectionType.All, "sheetPassword", null);

            // Verify that the protected range is still password‑protected
            Console.WriteLine($"After protecting sheet, range password protected: {protectedRange.IsProtectedWithPassword}");

            // Save the workbook
            string filePath = "ProtectedRangeDemo.xlsx";
            workbook.Save(filePath);

            // Load the saved workbook to verify persistence
            Workbook loadedWorkbook = new Workbook(filePath);
            Worksheet loadedWorksheet = loadedWorkbook.Worksheets[0];
            ProtectedRange loadedRange = loadedWorksheet.AllowEditRanges[0];

            // Output verification results
            Console.WriteLine($"Loaded worksheet is protected: {loadedWorksheet.IsProtected}");
            Console.WriteLine($"Loaded range password protected: {loadedRange.IsProtectedWithPassword}");
        }
    }
}
