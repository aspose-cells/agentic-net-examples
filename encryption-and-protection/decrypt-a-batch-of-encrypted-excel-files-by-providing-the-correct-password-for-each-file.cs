// Title: Batch decrypt multiple password‑protected Excel .xlsx files using Aspose.Cells for .NET
// AI Prompts: Create a C# console program that loops through a Dictionary<string,string> of encrypted Excel file paths and their passwords, opens each workbook with Aspose.Cells LoadOptions, clears the workbook password, and writes the unprotected file to a target directory. | Enhance the program to catch exceptions per file and write a CSV log containing the source path, decryption result, and error details while continuing the batch processing.
// Common Searches: asp.net batch decrypt password protected xlsx files using Aspose.Cells | c# load encrypted Excel workbook with password and save without password | how to remove workbook password from multiple Excel files programmatically | process a list of Excel files and passwords to create unprotected copies in C# | aspose.cells decrypt multiple workbooks in a loop
// Tags: multiple workbook decryption Aspose.Cells | open encrypted Excel file via LoadOptions | clear workbook protection Aspose.Cells | save unprotected xlsx Aspose.Cells | dictionary mapping file paths to passwords C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace ExcelDecryption
{
    // The example iterates over a dictionary of encrypted .xlsx file paths and their passwords, loads each workbook with Aspose.Cells using LoadOptions.Password, clears any workbook password, and saves an unprotected copy to a specified output folder, handling missing files and runtime exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            // Map of encrypted file paths to their passwords
            var filesWithPasswords = new Dictionary<string, string>
            {
                // Example entries – replace with actual paths and passwords
                { @"C:\Encrypted\Report1.xlsx", "Password123" },
                { @"C:\Encrypted\Report2.xlsx", "Secret!@#" },
                // Add more files as needed
            };

            // Destination folder for decrypted files
            string outputFolder = @"C:\Decrypted";
            Directory.CreateDirectory(outputFolder);

            foreach (var kvp in filesWithPasswords)
            {
                string encryptedPath = kvp.Key;
                string password = kvp.Value;

                // Verify source file exists
                if (!File.Exists(encryptedPath))
                {
                    Console.WriteLine($"Source file not found: {encryptedPath}");
                    continue;
                }

                try
                {
                    // Load the encrypted workbook using the password
                    var loadOptions = new LoadOptions(LoadFormat.Xlsx)
                    {
                        Password = password
                    };

                    var workbook = new Workbook(encryptedPath, loadOptions);

                    // Remove workbook password (if any)
                    workbook.Settings.Password = null;

                    // Build output file path
                    string fileName = Path.GetFileNameWithoutExtension(encryptedPath);
                    string decryptedPath = Path.Combine(outputFolder, $"{fileName}_decrypted.xlsx");

                    // Save the workbook without a password
                    workbook.Save(decryptedPath, SaveFormat.Xlsx);

                    Console.WriteLine($"Decrypted file saved to: {decryptedPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{encryptedPath}': {ex.Message}");
                }
            }

            Console.WriteLine("All files processed.");
        }
    }
}
