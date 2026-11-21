// Title: How to programmatically unprotect a password‑protected Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# code to remove worksheet protection with a password using Aspose.Cells | Aspose.Cells unprotect Excel file programmatically when password is known | Load a protected .xlsx workbook and save an unprotected copy with Aspose.Cells .NET | Handling invalid password exception while unprotecting an Excel sheet in C# | Unprotect the first worksheet of an Excel workbook using the Aspose.Cells API
// Tags: Aspose.Cells worksheet unprotect C# | remove Excel sheet password Aspose.Cells | load protected workbook Aspose.Cells | save unprotected Excel file .NET | CellsException invalid password handling

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads 'ProtectedTable.xlsx' with Aspose.Cells, calls Worksheet.Unprotect using the supplied password to lift protection from the first worksheet, catches a CellsException if the password is incorrect, and then saves the workbook as 'UnprotectedTable.xlsx' without protection.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "ProtectedTable.xlsx";
            string outputPath = "UnprotectedTable.xlsx";
            string password = "YourPassword"; // replace with actual password

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                try
                {
                    // Unprotect the first worksheet using the provided password
                    Worksheet sheet = workbook.Worksheets[0];
                    sheet.Unprotect(password);
                }
                catch (CellsException ex)
                {
                    // Aspose.Cells throws CellsException when the password is invalid
                    Console.WriteLine("Invalid password provided for unprotecting the worksheet.");
                    Console.WriteLine($"Details: {ex.Message}");
                    return;
                }

                // Save the unprotected workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved as {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
