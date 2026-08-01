// Title: C# – Batch encrypt Excel files in a folder with Aspose.Cells using filename‑based passwords
// Description: A console app that scans a specified directory, selects Excel workbooks (.xlsx, .xls, .xlsb, .xlsm), loads each with Aspose.Cells, assigns a password derived from the file name (without extension), applies 128‑bit strong encryption, overwrites the original file, and writes success or error messages to the console.
// Keywords: Aspose.Cells C# encrypt Excel | batch Excel password protection .NET | filename based workbook password | strong encryption Excel files | programmatic Excel security C# | encrypt multiple Excel files | Aspose.Cells encryption example
// Common Searches: how to encrypt all Excel files in a folder using Aspose.Cells | C# batch set password for Excel workbooks | Aspose.Cells encrypt multiple .xlsx files | apply strong encryption to Excel files programmatically | set workbook password from file name C#
// Developer Intent: Automatically protect every Excel workbook in a directory by assigning a unique password derived from its file name.
// Use Cases: Secure a collection of financial spreadsheets before archiving, giving each file its own name‑based password. | Enforce per‑report encryption for user‑generated Excel outputs to satisfy compliance policies. | Prepare a set of Excel templates for distribution, locking each with a distinct password without manual effort.
// AI Prompts: Write C# code that uses Aspose.Cells to encrypt all .xlsx, .xls, .xlsb, and .xlsm files in a given folder, using the file name (without extension) as the password and applying 128‑bit strong encryption. | Refactor the batch encryption program to log progress and errors to a file and to skip files that are already password‑protected. | Explain how to modify the example so that the password is a SHA‑256 hash of the file name instead of the plain name, while still using Aspose.Cells encryption.

using System;
using System.IO;
using Aspose.Cells;

namespace BatchEncryptExcel
{
    // A console app that scans a specified directory, selects Excel workbooks (.xlsx, .xls, .xlsb, .xlsm), loads each with Aspose.Cells, assigns a password derived from the file name (without extension), applies 128‑bit strong encryption, overwrites the original file, and writes success or error messages to the console.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the Excel files to encrypt
            string folderPath = @"C:\ExcelFolder";

            // Verify that the folder exists
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder not found: {folderPath}");
                return;
            }

            try
            {
                // Get all files in the folder (filter later by extension)
                string[] allFiles = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly);

                foreach (string filePath in allFiles)
                {
                    string extension = Path.GetExtension(filePath).ToLowerInvariant();
                    if (extension != ".xlsx" && extension != ".xls" && extension != ".xlsb" && extension != ".xlsm")
                        continue; // Skip non‑Excel files

                    // Ensure the file still exists before loading
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found (skipped): {filePath}");
                        continue;
                    }

                    try
                    {
                        // Load the workbook
                        Workbook workbook = new Workbook(filePath);

                        // Derive a unique password from the file name (without extension)
                        string password = Path.GetFileNameWithoutExtension(filePath);

                        // Set the password for the workbook (this encrypts the file)
                        workbook.Settings.Password = password;

                        // Optional: set stronger encryption options (ignored for .xlsx/.xlsm/.xlsb but kept for completeness)
                        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

                        // Save the workbook back to the same file (overwrites the original)
                        workbook.Save(filePath);
                        Console.WriteLine($"Encrypted: {Path.GetFileName(filePath)}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                    }
                }

                Console.WriteLine("Batch encryption completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
