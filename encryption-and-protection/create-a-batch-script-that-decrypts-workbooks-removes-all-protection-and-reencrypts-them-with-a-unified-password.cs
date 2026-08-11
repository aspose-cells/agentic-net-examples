// Title: C# batch script to decrypt, strip protection, and re‑encrypt Excel/ODS workbooks with a single password using Aspose.Cells
// Description: Scans a folder for Excel and ODS files, loads each workbook with its original password (if known), removes workbook, shared‑workbook and worksheet protection, optionally clears macros, digital signatures and personal information, then applies a unified strong encryption password and saves the files to an output directory.
// Keywords: Aspose.Cells | C# batch workbook processing | Excel decryption | remove worksheet protection | re‑encrypt workbooks | unified password | strong encryption AES | load options password | shared workbook unprotect | ODS encryption | .NET Excel security
// Common Searches: batch remove protection from Excel files Aspose.Cells | C# re‑encrypt multiple workbooks with one password | how to decrypt and re‑encrypt ODS files programmatically | Aspose.Cells remove shared workbook protection in bulk | set AES encryption for Excel files using Aspose.Cells
// Developer Intent: Open each workbook, clear all existing protections, and save it encrypted with a common password.
// Use Cases: Standardize password protection across a legacy collection of spreadsheets before distribution. | Clean up shared workbooks by removing shared mode, macros, and personal data prior to archiving. | Prepare a batch of confidential spreadsheets for secure sharing by applying a single strong password.
// AI Prompts: Generate C# code that uses Aspose.Cells to batch decrypt Excel/ODS files, remove all protection types, and re‑encrypt them with a specified password. | Explain how to handle workbooks when the original password is unknown or missing while using Aspose.Cells. | Show how to configure AES‑128 encryption options when saving workbooks with Aspose.Cells.

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;

namespace WorkbookBatchReencrypt
{
    // Scans a folder for Excel and ODS files, loads each workbook with its original password (if known), removes workbook, shared‑workbook and worksheet protection, optionally clears macros, digital signatures and personal information, then applies a unified strong encryption password and saves the files to an output directory.
    class Program
    {
        // Unified password to be applied to all processed workbooks
        private const string UnifiedPassword = "UnifiedPass123";

        static void Main(string[] args)
        {
            // Folder containing the workbooks to process
            string inputFolder = @"C:\InputWorkbooks";
            // Folder where the re‑encrypted workbooks will be saved
            string outputFolder = @"C:\OutputWorkbooks";

            // Mapping of workbook file name to its current password (if known)
            // If a workbook is not password‑protected, leave the value null or empty
            var originalPasswords = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                // Example entries:
                // { "Encrypted1.xlsx", "oldPass1" },
                // { "Encrypted2.xlsm", "oldPass2" }
            };

            // Ensure output directory exists
            Directory.CreateDirectory(outputFolder);

            // Process each supported workbook file in the input folder
            foreach (string filePath in Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly))
            {
                string extension = Path.GetExtension(filePath).ToLowerInvariant();
                if (extension != ".xlsx" && extension != ".xls" && extension != ".xlsm" && extension != ".ods")
                    continue; // Skip unsupported files

                string fileName = Path.GetFileName(filePath);
                Console.WriteLine($"Processing: {fileName}");

                // Determine the original password (if any) for this file
                originalPasswords.TryGetValue(fileName, out string originalPassword);

                // Load the workbook (with password if it is encrypted)
                LoadOptions loadOptions = new LoadOptions();
                if (!string.IsNullOrEmpty(originalPassword))
                    loadOptions.Password = originalPassword;

                Workbook workbook = new Workbook(filePath, loadOptions);

                // ----- Remove workbook‑level protection -----
                if (workbook.IsWorkbookProtectedWithPassword)
                {
                    // Unprotect using the original password (empty string if none)
                    workbook.Unprotect(originalPassword ?? string.Empty);
                }

                // ----- Remove shared workbook protection (if any) -----
                try
                {
                    workbook.UnprotectSharedWorkbook(originalPassword ?? string.Empty);
                }
                catch
                {
                    // Ignored – workbook may not be a shared workbook
                }

                // ----- Remove worksheet protection for all worksheets -----
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    if (sheet.IsProtected)
                    {
                        sheet.Unprotect(originalPassword ?? string.Empty);
                    }
                }

                // ----- Optional cleanup (macros, digital signatures, personal info) -----
                try { workbook.RemoveMacro(); } catch { }
                try { workbook.RemoveDigitalSignature(); } catch { }
                try { workbook.RemovePersonalInformation(); } catch { }

                // ----- Apply unified encryption password -----
                workbook.Settings.Password = UnifiedPassword;

                // Set strong encryption options (optional but recommended)
                workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

                // ----- Save the re‑encrypted workbook -----
                string outputPath = Path.Combine(outputFolder, fileName);
                workbook.Save(outputPath);

                Console.WriteLine($"Saved re‑encrypted workbook to: {outputPath}");
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}
