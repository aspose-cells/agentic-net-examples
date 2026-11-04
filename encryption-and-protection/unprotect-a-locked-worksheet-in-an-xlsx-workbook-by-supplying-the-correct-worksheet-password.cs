// Title: How to unprotect a password‑protected worksheet in an XLSX file using Aspose.Cells for .NET (C#)
// Common Searches: C# code to remove worksheet protection with Aspose.Cells | Aspose.Cells unprotect Excel sheet using worksheet password | Load password‑protected workbook and unlock sheet in .NET | Programmatically unprotect a locked worksheet in XLSX with Aspose.Cells | Aspose.Cells example for the Worksheet.Unprotect method in C#
// Tags: Aspose.Cells worksheet unprotect C# | load password protected Excel workbook Aspose.Cells | remove worksheet password .NET | save unprotected XLSX Aspose.Cells | unprotect sheet using Unprotect method

using System;
using System.IO;
using Aspose.Cells;

// The example loads an XLSX workbook (optionally providing a workbook password), accesses the target worksheet, calls Worksheet.Unprotect with the correct worksheet password to remove protection, and then saves the workbook as an unprotected file.
class UnprotectWorksheetExample
{
    static void Main()
    {
        // Path to the source XLSX file (locked worksheet)
        string inputPath = "input.xlsx";

        // Password used to protect the worksheet
        string worksheetPassword = "myPassword";

        // Password used to protect the workbook file itself (if any)
        string workbookPassword = ""; // set if the workbook is password‑protected

        // Path where the unprotected workbook will be saved
        string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook, providing the workbook password if required
            LoadOptions loadOptions = new LoadOptions();
            if (!string.IsNullOrEmpty(workbookPassword))
                loadOptions.Password = workbookPassword;

            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Access the worksheet you want to unprotect (e.g., the first worksheet)
            Worksheet sheet = workbook.Worksheets[0];

            // Unprotect the worksheet using the correct password
            sheet.Unprotect(worksheetPassword);

            // Save the workbook with the worksheet now unprotected
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
