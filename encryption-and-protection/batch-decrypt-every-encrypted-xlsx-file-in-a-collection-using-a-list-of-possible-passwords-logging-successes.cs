// Title: Batch Decrypt Encrypted XLSX Files with Aspose.Cells for .NET
// Description: A C# console utility that scans a folder for password‑protected XLSX workbooks, detects encryption, tries a predefined list of passwords, removes the protection with Aspose.Cells, saves the unencrypted copy, and logs each successful decryption.
// Keywords: Aspose.Cells | C# | .NET | batch decrypt Excel | encrypted XLSX | password list | verify Excel password | remove workbook protection | detect encrypted file | automate Excel decryption
// Common Searches: how to batch decrypt XLSX files using Aspose.Cells | c# program to try multiple passwords on encrypted Excel workbooks | aspnet automate removal of Excel file password | verify Excel password before loading with Aspose.Cells | bulk unlock password‑protected Excel files .NET
// Developer Intent: Automatically unlock every encrypted XLSX file in a directory by testing a set of possible passwords and saving the decrypted versions.
// Use Cases: Process nightly drops of password‑protected reports so downstream analytics can read them. | Migrate a legacy archive of secured Excel files to an unprotected repository for easier access. | Identify which known password opens each workbook before performing data extraction or validation.
// AI Prompts: Generate C# code that uses Aspose.Cells to batch decrypt XLSX files with a supplied password list and produce a log of successes and failures. | Show how to extend the utility to export a CSV summary containing file name, successful password, and output path. | Suggest performance optimizations for decrypting large encrypted workbooks in bulk with Aspose.Cells.

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;

namespace BatchDecrypt
{
    // A C# console utility that scans a folder for password‑protected XLSX workbooks, detects encryption, tries a predefined list of passwords, removes the protection with Aspose.Cells, saves the unencrypted copy, and logs each successful decryption.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing encrypted XLSX files
            string inputFolder = @"EncryptedFiles";
            // Folder where decrypted files will be saved
            string outputFolder = @"DecryptedFiles";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Verify input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder '{inputFolder}' does not exist. Please create it and add encrypted files.");
                return;
            }

            // List of possible passwords to try
            List<string> possiblePasswords = new List<string>
            {
                "password1",
                "password2",
                "1234",
                "test"
            };

            // Process each .xlsx file in the input folder
            foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
            {
                try
                {
                    // Detect file format and check if the file is encrypted
                    FileFormatInfo fileInfo = FileFormatUtil.DetectFileFormat(filePath);
                    if (!fileInfo.IsEncrypted)
                    {
                        Console.WriteLine($"{Path.GetFileName(filePath)} is not encrypted. Skipping.");
                        continue;
                    }

                    bool decrypted = false;

                    // Try each password until one succeeds
                    foreach (string pwd in possiblePasswords)
                    {
                        // Verify password without loading the whole workbook
                        using (Stream stream = File.OpenRead(filePath))
                        {
                            if (FileFormatUtil.VerifyPassword(stream, pwd))
                            {
                                // Load the workbook with the correct password
                                LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto)
                                {
                                    Password = pwd
                                };
                                Workbook workbook = new Workbook(filePath, loadOptions);

                                // Remove the password protection
                                workbook.Settings.Password = null;

                                // Save the unprotected workbook
                                string outputPath = Path.Combine(
                                    outputFolder,
                                    Path.GetFileNameWithoutExtension(filePath) + "_decrypted.xlsx");

                                workbook.Save(outputPath);

                                Console.WriteLine($"Successfully decrypted '{Path.GetFileName(filePath)}' with password '{pwd}'. Saved to '{outputPath}'.");
                                decrypted = true;
                                break; // Exit password loop for this file
                            }
                        }
                    }

                    if (!decrypted)
                    {
                        Console.WriteLine($"Failed to decrypt '{Path.GetFileName(filePath)}'. No matching password found.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{Path.GetFileName(filePath)}': {ex.Message}");
                }
            }
        }
    }
}
