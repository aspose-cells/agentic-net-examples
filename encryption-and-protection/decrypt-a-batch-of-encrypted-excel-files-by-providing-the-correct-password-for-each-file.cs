// Title: C# Batch Decrypt Encrypted Excel Files with Aspose.Cells – Per‑File Passwords
// Description: A complete C# utility that scans a folder, detects encrypted Excel workbooks, opens each file with its specific password from a dictionary, removes the protection, and saves the decrypted copy to a target directory. Unprotected files are copied unchanged. Uses Aspose.Cells for .NET to handle .xls, .xlsx, .xlsm and .ods formats.
// Keywords: Aspose.Cells batch decrypt | C# decrypt encrypted Excel | remove Excel password programmatically | load password‑protected workbook Aspose | detect encrypted Excel file .NET | bulk Excel decryption C# | per‑file password Excel decryption | Aspose.Cells file format detection
// Common Searches: batch decrypt password protected Excel files C# | Aspose.Cells decrypt multiple .xlsx files with different passwords | how to detect encrypted Excel workbook before opening | C# script to copy unencrypted Excel files and decrypt protected ones | remove protection from a folder of Excel files using Aspose
// Developer Intent: Automatically open each encrypted Excel workbook with its assigned password, strip the workbook protection, and write the decrypted file to a separate folder while leaving non‑encrypted files untouched.
// Use Cases: Process nightly financial reports that arrive password‑protected, decrypt them for downstream analytics, and archive the clear copies. | Migrate a legacy collection of secured spreadsheets to an unprotected repository for migration to a data warehouse. | Create a pre‑processing step in an ETL pipeline that normalizes a mixed set of encrypted and plain Excel files before further transformation.
// AI Prompts: Write C# code using Aspose.Cells to batch decrypt Excel files from a folder, using a dictionary of filenames and passwords, and log results to a CSV file. | Show how to extend the utility to retry opening a workbook with an alternate password list when the first password fails. | Generate a PowerShell wrapper that calls the C# batch decryption executable and passes the source and destination paths as parameters.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBatchDecrypt
{
    // A complete C# utility that scans a folder, detects encrypted Excel workbooks, opens each file with its specific password from a dictionary, removes the protection, and saves the decrypted copy to a target directory. Unprotected files are copied unchanged. Uses Aspose.Cells for .NET to handle .xls, .xlsx, .xlsm and .ods formats.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Folder containing encrypted Excel files
                string sourceFolder = @"C:\EncryptedFiles";

                // Folder where decrypted files will be saved
                string outputFolder = @"C:\DecryptedFiles";

                // Mapping of file names (or full paths) to their passwords
                var passwordMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    // Example entries
                    // { "Report1.xlsx", "Password123" },
                    // { "FinanceData.xls", "Secret!" }
                };

                // Verify source folder exists before proceeding
                if (!Directory.Exists(sourceFolder))
                {
                    Console.WriteLine($"Source folder does not exist: {sourceFolder}");
                    return;
                }

                DecryptBatch(sourceFolder, outputFolder, passwordMap);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        /// <param name="sourceFolder">Folder containing the encrypted files.</param>
        /// <param name="outputFolder">Folder where decrypted files will be written.</param>
        /// <param name="passwordMap">Dictionary that maps a file name to its password.</param>
        static void DecryptBatch(string sourceFolder, string outputFolder, Dictionary<string, string> passwordMap)
        {
            // Ensure output directory exists
            Directory.CreateDirectory(outputFolder);

            // Get all Excel files (common extensions) in the source folder
            string[] excelFiles = Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly);
            foreach (string filePath in excelFiles)
            {
                string extension = Path.GetExtension(filePath).ToLowerInvariant();
                if (extension != ".xls" && extension != ".xlsx" && extension != ".xlsm" && extension != ".ods")
                    continue; // Skip non‑Excel files

                // Detect file format and encryption status
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
                Console.WriteLine($"Processing '{Path.GetFileName(filePath)}' - Encrypted: {formatInfo.IsEncrypted}");

                if (!formatInfo.IsEncrypted)
                {
                    // If not encrypted, simply copy the file to the output folder
                    string destPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
                    try
                    {
                        File.Copy(filePath, destPath, overwrite: true);
                        Console.WriteLine($"  Copied unencrypted file to '{destPath}'.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"  Failed to copy file: {ex.Message}");
                    }
                    continue;
                }

                // Retrieve the password for this file
                if (!passwordMap.TryGetValue(Path.GetFileName(filePath), out string password))
                {
                    Console.WriteLine($"  No password supplied for '{Path.GetFileName(filePath)}'. Skipping.");
                    continue;
                }

                // Verify the file still exists before attempting to load
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"  File not found: {filePath}. Skipping.");
                    continue;
                }

                // Load the workbook with the password
                var loadOptions = new LoadOptions(LoadFormat.Auto) { Password = password };
                Workbook workbook;
                try
                {
                    workbook = new Workbook(filePath, loadOptions);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"  Failed to open '{Path.GetFileName(filePath)}' with provided password: {ex.Message}");
                    continue;
                }

                // Remove encryption by clearing the password setting
                workbook.Settings.Password = null; // Null or empty removes protection

                // Save the decrypted workbook
                string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"  Decrypted and saved to '{outputPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"  Failed to save decrypted file: {ex.Message}");
                }
            }
        }
    }
}
