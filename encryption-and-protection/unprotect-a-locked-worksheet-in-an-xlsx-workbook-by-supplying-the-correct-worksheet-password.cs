// Title: How to Unprotect a Password‑Protected Worksheet in an XLSX File Using Aspose.Cells for .NET
// Description: Load a protected XLSX workbook with Aspose.Cells, call Worksheet.Unprotect with the correct password, confirm the sheet is no longer locked via the IsProtected flag, and save the workbook as an unprotected file.
// Keywords: Aspose.Cells unprotect worksheet C# | remove Excel sheet password .NET | Worksheet.Unprotect method | Excel protection removal programmatically | C# unprotect Excel worksheet | IsProtected property Aspose.Cells | save unprotected workbook Aspose | decrypt Excel sheet using Aspose.Cells
// Common Searches: Aspose.Cells unprotect worksheet example | C# remove password from Excel sheet | Worksheet.Unprotect usage in .NET | check if Excel sheet is protected after Unprotect | save Excel file after removing sheet protection
// Developer Intent: The developer needs to programmatically remove password protection from a specific worksheet in an XLSX workbook using Aspose.Cells for .NET.
// Use Cases: A one‑off utility that opens a workbook, unprotects the first sheet with a known password, and writes the result to a new file. | A reusable method that accepts a file path, worksheet index or name, and password, then returns success status after unprotecting the sheet. | Batch processing of multiple workbooks where each worksheet is unlocked with its respective password before further data extraction.
// AI Prompts: Generate C# code that loads an XLSX file with Aspose.Cells, unprotects a given worksheet using a supplied password, verifies the protection state, and saves the unprotected workbook. | Create a .NET function that takes a Workbook object, worksheet name, and password, calls Worksheet.Unprotect, and returns the updated IsProtected value. | Show error‑handling patterns for catching an incorrect password exception when using Worksheet.Unprotect in Aspose.Cells.

using System;
using Aspose.Cells;

// Load a protected XLSX workbook with Aspose.Cells, call Worksheet.Unprotect with the correct password, confirm the sheet is no longer locked via the IsProtected flag, and save the workbook as an unprotected file.
class UnprotectWorksheetDemo
{
    static void Main()
    {
        // Load the workbook that contains a protected worksheet
        string inputFile = "protected.xlsx";
        Workbook workbook = new Workbook(inputFile);

        // Get the first worksheet (index 0)
        Worksheet worksheet = workbook.Worksheets[0];

        // Supply the correct password to unprotect the worksheet
        string password = "myPassword";
        worksheet.Unprotect(password);

        // Verify that the worksheet is now unprotected
        Console.WriteLine("Worksheet is protected: " + worksheet.IsProtected);

        // Save the workbook with the worksheet now unprotected
        string outputFile = "unprotected.xlsx";
        workbook.Save(outputFile);
    }
}
