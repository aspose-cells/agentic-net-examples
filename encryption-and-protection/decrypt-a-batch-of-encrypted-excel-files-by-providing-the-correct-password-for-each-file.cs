using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;

class DecryptExcelBatch
{
    static void Main()
    {
        // Folder containing encrypted Excel files
        string inputFolder = @"C:\EncryptedFiles";
        // Folder where decrypted files will be saved
        string outputFolder = @"C:\DecryptedFiles";

        // Verify input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder not found: {inputFolder}");
            return;
        }

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Map each file name to its password (populate as needed)
        var passwordMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "file1.xlsx", "password1" },
            { "file2.xls",  "password2" }
            // Add more entries for other files
        };

        // Process each file in the input folder
        foreach (string filePath in Directory.GetFiles(inputFolder))
        {
            string fileName = Path.GetFileName(filePath);
            Console.WriteLine($"Processing: {fileName}");

            try
            {
                // Detect format and encryption status
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
                Console.WriteLine($"  Format: {formatInfo.FileFormatType}, Encrypted: {formatInfo.IsEncrypted}");

                // If the file is not encrypted, copy it unchanged
                if (!formatInfo.IsEncrypted)
                {
                    string destPath = Path.Combine(outputFolder, fileName);
                    File.Copy(filePath, destPath, true);
                    Console.WriteLine("  Not encrypted – copied without changes.");
                    continue;
                }

                // Retrieve the password for this file
                if (!passwordMap.TryGetValue(fileName, out string password))
                {
                    Console.WriteLine("  No password found – skipping file.");
                    continue;
                }

                // Ensure the file still exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine("  File not found – skipping.");
                    continue;
                }

                // Load the workbook using the password
                var loadOptions = new LoadOptions(LoadFormat.Auto) { Password = password };
                var workbook = new Workbook(filePath, loadOptions);

                // If the workbook itself has a protection password, unprotect it
                if (!string.IsNullOrEmpty(workbook.Settings.Password))
                {
                    workbook.Unprotect(password);
                }

                // Save the workbook without a password (decrypted)
                string outputPath = Path.Combine(outputFolder, fileName);
                workbook.Save(outputPath);
                Console.WriteLine($"  Decrypted and saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  Error processing {fileName}: {ex.Message}");
            }
        }

        Console.WriteLine("Batch decryption completed.");
    }
}