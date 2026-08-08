// Title: C# – Unprotect Worksheet, Copy Rows, and Re‑protect with Same Password using Aspose.Cells
// Description: A concise Aspose.Cells for .NET example that loads an Excel file, removes protection from the first worksheet with a known password, copies a block of rows using Cells.CopyRows, reapplies full protection with the same password, and saves the result.
// Keywords: Aspose.Cells C# | unprotect worksheet | protect worksheet password | copy rows Aspose.Cells | Cells.CopyRows example | Excel automation C# | worksheet protection API | row duplication code | GitHub Aspose.Cells sample | Excel file processing
// Common Searches: Aspose.Cells copy rows in protected sheet C# | how to unprotect and protect Excel worksheet with password using Aspose.Cells | C# example for Cells.CopyRows after unprotecting sheet | re‑apply worksheet protection after modifying rows Aspose.Cells | Aspose.Cells unprotect worksheet programmatically
// Developer Intent: Temporarily lift worksheet protection, duplicate specific rows, then restore the original password‑protected state in a .NET application.
// Use Cases: Replicate header rows in a locked template before inserting new data. | Create a summary section by copying data rows in a secured financial report. | Automate bulk import that requires row duplication while maintaining worksheet security.
// AI Prompts: Generate C# code with Aspose.Cells that unprotects a sheet, copies rows 5‑9 to row 15, and protects the sheet again using the original password. | Explain best practices for handling exceptions when calling Unprotect and Protect methods in Aspose.Cells. | Show how to copy rows from a protected worksheet to another worksheet while preserving the original protection settings.

using System;
using Aspose.Cells;

namespace AsposeCellsRowCopyExample
{
    // A concise Aspose.Cells for .NET example that loads an Excel file, removes protection from the first worksheet with a known password, copies a block of rows using Cells.CopyRows, reapplies full protection with the same password, and saves the result.
    class Program
    {
        static void Main()
        {
            // Input and output file paths
            string inputFile = "input.xlsx";
            string outputFile = "output.xlsx";

            // Worksheet protection password
            string password = "myPassword";

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputFile);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Unprotect the worksheet using the existing password (rule: Unprotect(string))
            sheet.Unprotect(password);

            // Copy rows within the same worksheet.
            // Example: copy rows 0‑4 (first five rows) to start at row 10.
            // Parameters: source cells, source start row, destination start row, number of rows.
            sheet.Cells.CopyRows(sheet.Cells, 0, 10, 5);

            // Protect the worksheet again with the same password (rule: Protect(ProtectionType, string, string))
            sheet.Protect(ProtectionType.All, password, null);

            // Save the modified workbook (lifecycle rule: save)
            workbook.Save(outputFile);
        }
    }
}
