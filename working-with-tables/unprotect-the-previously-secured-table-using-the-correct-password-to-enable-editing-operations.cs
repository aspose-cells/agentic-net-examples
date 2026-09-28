// Title: How to unprotect a password‑protected Excel table using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx workbook with Aspose.Cells, calls Worksheet.Unprotect with a supplied password to remove protection from the sheet and its tables, and saves the result as a new file. | Show an example of editing a previously protected list object after calling Worksheet.Unprotect in Aspose.Cells, including loading, unprotecting, modifying cells, and saving the workbook.
// Common Searches: asp.net aspose.cells unprotect worksheet with password c# example | remove protection from Excel table (list object) using Aspose.Cells .NET | c# code to edit a protected Excel table after unprotecting with Aspose.Cells | how to call Worksheet.Unprotect to unlock tables in an .xlsx file | aspose.cells unprotect protected list object and save workbook
// Tags: unprotect worksheet Aspose.Cells C# | remove table protection Aspose.Cells .NET | worksheet.unprotect password example | edit unprotected Excel list object Aspose.Cells | save unprotected workbook Aspose.Cells

using Aspose.Cells;

// Loads a password‑protected workbook, uses Worksheet.Unprotect with the known password to clear protection from the sheet and any embedded tables, then saves the workbook as an unprotected file ready for further editing.
class UnprotectTableExample
{
    static void Main()
    {
        // Load the existing workbook that contains the protected table
        Workbook workbook = new Workbook("ProtectedTable.xlsx");

        // Get the worksheet that holds the table (assuming the first worksheet)
        Worksheet sheet = workbook.Worksheets[0];

        // Unprotect the worksheet using the known password
        // This also removes protection from any tables (list objects) on the sheet
        sheet.Unprotect("YourPasswordHere");

        // At this point the table can be edited (e.g., rows added, cells modified, etc.)

        // Save the workbook after unprotecting
        workbook.Save("UnprotectedTable.xlsx");
    }
}
