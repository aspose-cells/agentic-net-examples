// Title: C# – Unprotect an Excel worksheet with a password using Aspose.Cells for .NET
// Description: Loads "ProtectedWorkbook.xlsx", accesses the first worksheet, removes its password protection with Worksheet.Unprotect("mySecretPassword"), and saves the result as "UnprotectedWorkbook.xlsx". Includes file‑existence check and exception handling.
// Keywords: Aspose.Cells unprotect worksheet C# | remove Excel worksheet password .NET | Worksheet.Unprotect example | C# Aspose.Cells unprotect workbook | programmatically unprotect Excel file
// Common Searches: how to unprotect a worksheet using Aspose.Cells C# | Aspose.Cells remove password protection from Excel sheet | C# code to unprotect protected Excel workbook | unprotect Excel worksheet programmatically Aspose
// Developer Intent: Remove password protection from a worksheet so it can be edited or processed further.
// Use Cases: Open a password‑protected workbook, unprotect a specific sheet, and save the file for downstream data manipulation. | Batch‑process multiple protected Excel files by iterating worksheets, applying Worksheet.Unprotect with known passwords, and exporting unprotected copies. | Integrate worksheet unprotection into an automated reporting pipeline that updates cell values after the protection is removed.
// AI Prompts: Generate C# code that opens a protected Excel file, unprotects the first worksheet using the password 'mySecretPassword', and saves the unprotected version. | Create a reusable method in C# that accepts input path, password, and output path, then uses Aspose.Cells to remove worksheet protection with robust error handling. | Show how to unprotect all worksheets in a workbook using Aspose.Cells for .NET, including optional password handling for each sheet.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads "ProtectedWorkbook.xlsx", accesses the first worksheet, removes its password protection with Worksheet.Unprotect("mySecretPassword"), and saves the result as "UnprotectedWorkbook.xlsx". Includes file‑existence check and exception handling.
    public class UnprotectWorksheetDemo
    {
        public static void Run()
        {
            string inputPath = "ProtectedWorkbook.xlsx";
            string outputPath = "UnprotectedWorkbook.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook that contains the protected worksheet
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Unprotect the worksheet using the correct password
                sheet.Unprotect("mySecretPassword");

                // Save the workbook after unprotection
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved as {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            UnprotectWorksheetDemo.Run();
        }
    }
}
