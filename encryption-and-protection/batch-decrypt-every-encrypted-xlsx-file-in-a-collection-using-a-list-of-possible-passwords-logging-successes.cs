// Title: Batch decrypt password‑protected XLSX files in C# using Aspose.Cells with a list of possible passwords
// AI Prompts: Write a C# console application that scans a source directory for *.xlsx files, attempts to open each workbook with every password from a predefined list using Aspose.Cells LoadOptions, and saves the successfully opened workbook to a destination folder without a password. | Enhance the batch decryption program to create a CSV log that records the original file name, the password that succeeded, and the path of the decrypted file for each successful operation. | Extend the solution to handle both .xlsx and .xls files by detecting the file extension and selecting the appropriate LoadFormat when constructing LoadOptions.
// Common Searches: asp.net batch decrypt multiple encrypted Excel workbooks using Aspose.Cells and password list | c# try several passwords on password‑protected xlsx with Aspose.Cells LoadOptions | how to remove password protection from a folder of Excel files programmatically in .NET | aspocells decrypt xlsx files automatically using a list of possible passwords | log successful Excel decryption results to CSV with Aspose.Cells C#
// Tags: batch decrypt password‑protected xlsx Aspose.Cells | load encrypted workbook with password list .NET | save workbook without password Aspose.Cells | iterate directory of encrypted Excel files C# | csv logging of decryption results Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace BatchDecryptXlsx
{
    // The program iterates over all XLSX files in a specified source folder, tries each password from a supplied list via Aspose.Cells LoadOptions, saves the workbook without a password to a target folder when decryption succeeds, and reports success, failure, or errors to the console.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing encrypted XLSX files
            string sourceFolder = @"C:\EncryptedFiles";

            // Folder where decrypted files will be saved
            string destFolder = @"C:\DecryptedFiles";

            // Verify source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"[ERROR] Source folder not found: {sourceFolder}");
                return;
            }

            // Ensure destination folder exists
            if (!Directory.Exists(destFolder))
            {
                Directory.CreateDirectory(destFolder);
            }

            // List of possible passwords to try
            List<string> possiblePasswords = new List<string>
            {
                "password1",
                "12345",
                "letmein",
                // add more passwords as needed
            };

            // Process each XLSX file in the source folder
            foreach (string filePath in Directory.GetFiles(sourceFolder, "*.xlsx"))
            {
                // Guard against missing file (should not happen, but safe)
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"[WARN] File not found: {filePath}");
                    continue;
                }

                bool decrypted = false;

                foreach (string pwd in possiblePasswords)
                {
                    try
                    {
                        // Load the workbook using the current password
                        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                        {
                            Password = pwd
                        };

                        Workbook workbook = new Workbook(filePath, loadOptions);

                        // If loading succeeds, the password is correct.
                        // Save the workbook without a password to remove encryption.
                        string fileNameWithoutExt = Path.GetFileNameWithoutExtension(filePath);
                        string destPath = Path.Combine(destFolder, $"{fileNameWithoutExt}_decrypted.xlsx");

                        workbook.Save(destPath, SaveFormat.Xlsx);

                        Console.WriteLine($"[SUCCESS] Decrypted '{Path.GetFileName(filePath)}' with password '{pwd}'. Saved to '{destPath}'.");
                        decrypted = true;
                        break; // Stop trying other passwords for this file
                    }
                    catch (CellsException)
                    {
                        // Incorrect password – continue with the next one
                    }
                    catch (Exception ex)
                    {
                        // Unexpected error – log and move to next file
                        Console.WriteLine($"[ERROR] Unexpected error processing '{Path.GetFileName(filePath)}': {ex.Message}");
                        break;
                    }
                }

                if (!decrypted)
                {
                    Console.WriteLine($"[FAIL] Could not decrypt '{Path.GetFileName(filePath)}' with provided passwords.");
                }
            }

            Console.WriteLine("Batch decryption completed.");
        }
    }
}
