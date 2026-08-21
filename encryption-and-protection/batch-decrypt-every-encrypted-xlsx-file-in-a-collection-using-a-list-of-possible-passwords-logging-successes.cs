// Title: C# Batch Decrypt Encrypted XLSX Files with Multiple Passwords Using Aspose.Cells
// Description: A C# console app that scans a directory (recursively), detects encrypted .xlsx workbooks, tries a predefined list of passwords, removes the password when a match is found, saves an unprotected copy with a "_decrypted" suffix, and logs success or failure for each file.
// Keywords: Aspose.Cells batch decryption | C# decrypt encrypted Excel | verify password Aspose.Cells | remove workbook password programmatically | detect encrypted XLSX files | load encrypted workbook with password | GitHub Aspose.Cells example | bulk Excel password removal
// Common Searches: batch decrypt encrypted Excel files C# | Aspose.Cells try multiple passwords | remove password from many XLSX files | C# script to unlock encrypted workbooks | detect and decrypt protected Excel files
// Developer Intent: Automatically unlock every encrypted XLSX file in a folder by testing a set of possible passwords and save each workbook without protection.
// Use Cases: Mass‑unprotect archived spreadsheets before migration to a data lake. | Process user‑submitted encrypted reports when the password list is known, producing plain‑text files for analysis. | Add a pre‑release check in CI/CD pipelines to ensure no password‑protected Excel files are shipped.
// AI Prompts: Write C# code with Aspose.Cells that recursively scans a folder, detects encrypted .xlsx files, attempts a list of passwords, removes the password, saves a decrypted copy, and logs each outcome. | Show how to modify the batch decryption script to generate a CSV report containing file path, successful password (if any), and status. | Explain error handling for files where no password matches and how to continue processing the remaining workbooks.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace BatchDecryptXlsx
{
    // A C# console app that scans a directory (recursively), detects encrypted .xlsx workbooks, tries a predefined list of passwords, removes the password when a match is found, saves an unprotected copy with a "_decrypted" suffix, and logs success or failure for each file.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the XLSX files to process
            string folderPath = @"C:\ExcelFiles";

            // List of possible passwords to try
            List<string> possiblePasswords = new List<string>
            {
                "password1",
                "password2",
                "1234",
                "test"
            };

            // Get all .xlsx files in the folder (including subfolders)
            string[] excelFiles = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.AllDirectories);

            foreach (string filePath in excelFiles)
            {
                // Detect if the file is encrypted
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
                if (!formatInfo.IsEncrypted)
                {
                    Console.WriteLine($"File is not encrypted, skipping: {filePath}");
                    continue;
                }

                bool decrypted = false;

                // Try each password until one succeeds
                foreach (string pwd in possiblePasswords)
                {
                    // Verify password using a fresh stream each time
                    using (FileStream stream = File.OpenRead(filePath))
                    {
                        if (FileFormatUtil.VerifyPassword(stream, pwd))
                        {
                            // Load the workbook with the correct password
                            LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
                            loadOptions.Password = pwd;
                            Workbook workbook = new Workbook(filePath, loadOptions);

                            // Remove the password protection
                            workbook.Settings.Password = null;

                            // Save the unprotected workbook (you can change the output path as needed)
                            string outputPath = Path.Combine(
                                Path.GetDirectoryName(filePath),
                                Path.GetFileNameWithoutExtension(filePath) + "_decrypted.xlsx");

                            workbook.Save(outputPath);
                            Console.WriteLine($"Successfully decrypted '{filePath}' with password '{pwd}'. Saved as '{outputPath}'.");
                            decrypted = true;
                            break; // Stop trying other passwords for this file
                        }
                    }
                }

                if (!decrypted)
                {
                    Console.WriteLine($"Failed to decrypt '{filePath}'. No matching password found.");
                }
            }
        }
    }
}
