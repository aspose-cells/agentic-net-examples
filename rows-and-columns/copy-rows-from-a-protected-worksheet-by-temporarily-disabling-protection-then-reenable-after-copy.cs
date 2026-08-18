// Title: Copy Rows in a Protected Worksheet with Aspose.Cells for .NET – Unprotect, Copy, Re‑protect
// Description: Load a workbook, detect if the first worksheet is locked, temporarily unprotect it with a known password, copy a range of rows using Cells.CopyRows, then restore full protection before saving the file.
// Keywords: Aspose.Cells copy rows protected sheet | C# unprotect worksheet Aspose.Cells | Cells.CopyRows example | reapply worksheet protection .NET | temporary sheet unprotect Aspose.Cells | protect type all Aspose.Cells | copy rows between indices C# | workbook manipulation Aspose.Cells | protected worksheet operations | Aspose.Cells .NET API
// Common Searches: how to copy rows from a protected sheet using Aspose.Cells | unprotect worksheet programmatically Aspose.Cells C# | Cells.CopyRows source and destination indices example | re‑apply protection after modifying a worksheet Aspose.Cells | copy rows in a locked Excel file with Aspose.Cells
// Developer Intent: Temporarily disable protection, copy rows, then restore protection.
// Use Cases: Duplicate header rows in a locked template while keeping the sheet secured. | Move or replicate data rows in a protected financial report before generating a new version. | Create a copy of specific rows in a password‑protected workbook for audit purposes without altering protection settings.
// AI Prompts: Generate C# code that uses Aspose.Cells to unprotect a worksheet, copy a set of rows with Cells.CopyRows, and re‑apply full protection. | Explain step‑by‑step how to handle password‑protected worksheets when copying rows in Aspose.Cells for .NET. | Create a reusable method that accepts source row range, destination index, and password to copy rows in a protected sheet using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsRowCopyProtected
{
    // Load a workbook, detect if the first worksheet is locked, temporarily unprotect it with a known password, copy a range of rows using Cells.CopyRows, then restore full protection before saving the file.
    class Program
    {
        static void Main()
        {
            // Path to the source workbook (must exist)
            string inputPath = "input.xlsx";
            // Path for the resulting workbook
            string outputPath = "output.xlsx";

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (assumed to be protected)
            Worksheet sheet = workbook.Worksheets[0];

            // Password used for protection (must match the existing one)
            string password = "pwd";

            // Temporarily remove protection if the sheet is protected
            if (sheet.IsProtected)
            {
                // Unprotect using the known password
                sheet.Unprotect(password);
            }

            // Example: copy the first three rows (0,1,2) to start at row index 5 (rows 5,6,7)
            // CopyRows(sourceCells, sourceRowIndex, destinationRowIndex, rowNumber)
            sheet.Cells.CopyRows(sheet.Cells, 0, 5, 3);

            // Re‑apply protection with the same password and all protection types
            sheet.Protect(ProtectionType.All, password, null);

            // Save the modified workbook
            workbook.Save(outputPath);
        }
    }
}
