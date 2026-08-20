// Title: Unprotect Excel workbook structure with a password using Aspose.Cells for .NET
// Description: Loads a password‑protected Excel file (protected_workbook.xlsx) with Aspose.Cells, removes the workbook‑structure protection via Workbook.Unprotect, and saves the result as an unprotected file (unprotected_workbook.xlsx). Includes checks for missing files, load failures, invalid passwords, and save errors.
// Keywords: Aspose.Cells unprotect workbook | C# remove workbook structure password | Workbook.Unprotect example | Excel protection removal .NET | save unprotected Excel file Aspose
// Common Searches: Aspose.Cells unprotect workbook structure C# | How to remove password from Excel workbook using .NET | Workbook.Unprotect with password example | C# code to open protected Excel file and save unprotected version | Aspose.Cells error handling for invalid password
// Developer Intent: Strip the workbook‑structure password from an Excel file and write the file back without protection.
// Use Cases: Automate bulk de‑protection of Excel workbooks before data extraction. | Integrate into a migration workflow where protected files must be opened, modified, and re‑saved. | Prepare a workbook for further programmatic changes (adding sheets, editing formulas) after removing structure protection.
// AI Prompts: Write C# code that opens a password‑protected Excel workbook with Aspose.Cells, calls Workbook.Unprotect using a supplied password, and saves the unprotected file with full error handling. | Show how to detect and report an incorrect password when calling Workbook.Unprotect in Aspose.Cells. | Explain how to unprotect only the workbook structure while keeping individual worksheet protections intact using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsUnprotectExample
{
    // Loads a password‑protected Excel file (protected_workbook.xlsx) with Aspose.Cells, removes the workbook‑structure protection via Workbook.Unprotect, and saves the result as an unprotected file (unprotected_workbook.xlsx). Includes checks for missing files, load failures, invalid passwords, and save errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the protected workbook
            string inputPath = "protected_workbook.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Password used to protect the workbook structure (if any)
            string password = "myPassword";

            Workbook workbook = null;

            try
            {
                // Load the protected workbook
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            try
            {
                // Attempt to unprotect the workbook using the provided password
                workbook.Unprotect(password);
            }
            catch (Exception ex)
            {
                // Aspose.Cells throws a generic exception when the password is invalid
                Console.WriteLine($"Failed to unprotect workbook: {ex.Message}");
                return;
            }

            // Save the unprotected workbook
            string outputPath = "unprotected_workbook.xlsx";

            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook has been unprotected and saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
