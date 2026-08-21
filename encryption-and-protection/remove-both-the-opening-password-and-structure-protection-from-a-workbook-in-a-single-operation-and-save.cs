// Title: Remove Opening Password and Structure Protection from an Excel Workbook with Aspose.Cells for .NET
// Description: Loads a password‑protected .xlsx using LoadOptions, calls Workbook.Unprotect to clear structure protection, empties Workbook.Settings.Password to drop the opening encryption, and saves a new unprotected file. Includes checks for missing input files and automatic creation of the output directory.
// Keywords: Aspose.Cells remove password | unprotect Excel workbook .NET | clear workbook structure protection | delete opening encryption Aspose | C# load password protected workbook | Workbook.Unprotect example | Workbook.Settings.Password empty | save unprotected Excel file | Aspose.Cells file handling | programmatic Excel decryption
// Common Searches: Aspose.Cells remove workbook password and protection | C# unprotect Excel file and clear opening password | How to delete structure protection with Aspose.Cells | Load encrypted .xlsx and save without password .NET | Batch remove Excel passwords using Aspose
// Developer Intent: Strip both the opening encryption password and any workbook structure protection from an Excel file and write the result as an unprotected workbook.
// Use Cases: Convert a secured template into a freely editable copy for downstream processing. | Automate de‑protection of multiple workbooks before importing data into a reporting system. | Integrate into a CI pipeline that validates and then releases password‑free Excel assets.
// AI Prompts: Generate C# code that opens a password‑protected Excel workbook with Aspose.Cells, removes its structure protection, clears the opening password, and saves an unprotected version. | Explain why assigning an empty string to Workbook.Settings.Password removes the encryption password in Aspose.Cells. | Create robust error‑handling logic for cases where the supplied password is wrong or the workbook is already unprotected.

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // Loads a password‑protected .xlsx using LoadOptions, calls Workbook.Unprotect to clear structure protection, empties Workbook.Settings.Password to drop the opening encryption, and saves a new unprotected file. Includes checks for missing input files and automatic creation of the output directory.
    class RemoveWorkbookProtection
    {
        static void Main()
        {
            try
            {
                // Path to the password‑protected workbook
                string inputPath = "protected_workbook.xlsx";

                // Verify the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Password used to open the workbook and to protect its structure
                string password = "myPassword";

                // Load the workbook with the opening password
                LoadOptions loadOptions = new LoadOptions
                {
                    Password = password
                };
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Remove workbook structure protection (if any)
                try
                {
                    workbook.Unprotect(password);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Structure unprotect failed (may be already unprotected): {ex.Message}");
                }

                // Remove the opening (encryption) password
                workbook.Settings.Password = string.Empty;

                // Save the unprotected workbook
                string outputPath = "unprotected_workbook.xlsx";

                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved without protection: {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
