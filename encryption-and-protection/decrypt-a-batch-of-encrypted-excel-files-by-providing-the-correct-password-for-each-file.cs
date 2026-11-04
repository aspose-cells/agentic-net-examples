// Title: Batch decrypt multiple password‑protected .xlsx files with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells batch remove password from encrypted Excel workbooks | How to decrypt several .xlsx files programmatically using Aspose.Cells | Load encrypted Excel file with password and save unprotected using Aspose.Cells .NET | Decrypt multiple Excel files with different passwords using a dictionary in C# | Aspose.Cells example for bulk Excel decryption in .NET
// Tags: batch excel decryption Aspose.Cells .NET | load encrypted workbook with password Aspose | remove workbook opening password C# | save unprotected xlsx Aspose.Cells | dictionary based password mapping for Excel decryption

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The ExcelDecryptor class iterates over a dictionary that maps encrypted .xlsx file paths to their passwords, loads each workbook with Aspose.Cells using LoadOptions, clears the opening password, and saves the decrypted workbook to a specified output folder, handling missing files and runtime errors.
class ExcelDecryptor
{
    // Decrypts a batch of encrypted Excel files.
    // filePasswords: key = encrypted file path, value = password for that file.
    // outputDirectory: folder where decrypted files will be saved (original names).
    public static void DecryptFiles(Dictionary<string, string> filePasswords, string outputDirectory)
    {
        // Ensure output directory exists
        if (!Directory.Exists(outputDirectory))
            Directory.CreateDirectory(outputDirectory);

        foreach (var kvp in filePasswords)
        {
            string encryptedPath = kvp.Key;
            string password = kvp.Value;

            try
            {
                // Verify the source file exists
                if (!File.Exists(encryptedPath))
                {
                    Console.WriteLine($"File not found: {encryptedPath}");
                    continue;
                }

                // Load the workbook with the provided password
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    Password = password
                };
                Workbook workbook = new Workbook(encryptedPath, loadOptions);

                // Remove any opening password
                workbook.Settings.Password = null;

                // Prepare output path
                string fileName = Path.GetFileName(encryptedPath);
                string decryptedPath = Path.Combine(outputDirectory, fileName);

                // Save the workbook without a password
                workbook.Save(decryptedPath, SaveFormat.Xlsx);
                Console.WriteLine($"Decrypted: {decryptedPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{encryptedPath}': {ex.Message}");
            }
        }
    }

    // Example usage
    static void Main()
    {
        // Map of encrypted files to their passwords
        var files = new Dictionary<string, string>
        {
            { @"C:\Encrypted\Report1.xlsx", "Pass123" },
            { @"C:\Encrypted\Report2.xlsx", "Secret!@#" },
            // Add more files as needed
        };

        string outputFolder = @"C:\Decrypted";

        DecryptFiles(files, outputFolder);

        Console.WriteLine("Decryption completed.");
    }
}
