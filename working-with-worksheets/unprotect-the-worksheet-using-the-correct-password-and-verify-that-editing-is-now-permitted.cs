// Title: Unprotect a password‑protected worksheet in an XLSX workbook with Aspose.Cells for .NET and confirm cell editability
// AI Prompts: Write C# code that opens a password‑protected XLSX file using Aspose.Cells, calls Worksheet.Unprotect with the supplied password, modifies a cell, and saves the workbook without protection. | Show how to verify that a worksheet has been successfully unprotected by writing a test value to a cell after invoking Unprotect in Aspose.Cells. | Demonstrate loading a protected Excel file, removing its worksheet protection, and exporting an unprotected copy using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# how to remove worksheet password protection from an XLSX file | load encrypted Excel workbook with password and save unprotected using Aspose.Cells | programmatically unprotect worksheet and edit cells with Aspose.Cells for .NET | verify worksheet is editable after calling Unprotect in Aspose.Cells C# | example code for Aspose.Cells Unprotect method with password
// Tags: Aspose.Cells worksheet Unprotect method C# | load password‑protected XLSX Aspose.Cells | remove worksheet protection programmatically Aspose | edit cell after worksheet unprotected Aspose.Cells | save unprotected workbook Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// // Loads a password‑protected XLSX workbook with Aspose.Cells, unprotects the first worksheet using the provided password, writes a test value to cell A1 to confirm editability, and saves the result as an unprotected file.
class Program
{
    static void Main()
    {
        const string inputFile = "protected.xlsx";
        const string outputFile = "unprotected.xlsx";
        const string password = "myPassword";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file not found: {inputFile}");
                return;
            }

            // Load the password‑protected workbook
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };
            Workbook workbook = new Workbook(inputFile, loadOptions);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Unprotect the worksheet (if it has its own protection password)
            sheet.Unprotect(password);

            // Write a test value to verify editing is allowed
            sheet.Cells["A1"].PutValue("Edited after unprotect");

            // Save the modified workbook
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved as {outputFile}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
