// Title: Decrypt an Encrypted XLSX and Save Unencrypted with Aspose.Cells for .NET
// Description: Detects whether an XLSX file is password‑protected, prompts for the required password, loads the workbook with LoadOptions, removes any workbook‑level protection, and saves a new file without encryption while keeping all formulas, styles, and data intact.
// Keywords: Aspose.Cells decrypt XLSX | remove Excel password .NET | detect encrypted workbook Aspose | LoadOptions.Password example | Workbook.Unprotect usage | preserve formulas when saving Excel | C# Excel decryption Aspose
// Common Searches: how to open encrypted xlsx with Aspose.Cells | remove password from Excel file using C# | detect if Excel workbook is encrypted before loading | save unprotected workbook after decryption Aspose | preserve formulas when converting encrypted Excel
// Developer Intent: Load a password‑protected XLSX, optionally unprotect the workbook, and write an unencrypted copy that retains all content and formatting.
// Use Cases: Automated batch conversion of secured reports to plain XLSX for downstream analytics. | Processing user‑uploaded encrypted spreadsheets, decrypting them, and storing the clean version for further manipulation. | Removing workbook‑level protection after validation to enable editing in other tools.
// AI Prompts: Generate C# code that uses Aspose.Cells to detect an encrypted Excel file, ask for the password, load it, unprotect the workbook if needed, and save an unencrypted copy preserving formulas and formatting. | Create a reusable method `DecryptWorkbook(string sourcePath, string destPath)` that handles missing files, encryption detection, password input, and saves the decrypted workbook. | Explain the interaction between `LoadOptions.Password` and `Workbook.Unprotect` for removing file encryption and workbook protection in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Detects whether an XLSX file is password‑protected, prompts for the required password, loads the workbook with LoadOptions, removes any workbook‑level protection, and saves a new file without encryption while keeping all formulas, styles, and data intact.
class DecryptExcel
{
    static void Main()
    {
        // Paths for the encrypted source file and the unencrypted destination file
        string sourcePath = "encrypted.xlsx";
        string destPath = "decrypted.xlsx";

        try
        {
            // Verify that the source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Detect file format and encryption status
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(sourcePath);
            Console.WriteLine($"Is encrypted: {formatInfo.IsEncrypted}");

            // Load the workbook with appropriate options
            Workbook workbook;
            if (formatInfo.IsEncrypted)
            {
                // Prompt for the password required to open the encrypted file
                Console.Write("Enter password for the encrypted workbook: ");
                string password = Console.ReadLine() ?? string.Empty;

                LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto)
                {
                    Password = password
                };

                // Load using the constructor that accepts a file path and LoadOptions
                workbook = new Workbook(sourcePath, loadOptions);
            }
            else
            {
                // No encryption, load normally
                workbook = new Workbook(sourcePath);
            }

            // If the workbook itself is protected with a password, remove it
            if (workbook.IsWorkbookProtectedWithPassword)
            {
                Console.Write("Enter password to unprotect the workbook (if any): ");
                string protectPassword = Console.ReadLine() ?? string.Empty;
                workbook.Unprotect(protectPassword);
            }

            // Save the workbook without any encryption settings
            workbook.Save(destPath);
            Console.WriteLine($"Decrypted file saved to: {destPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
