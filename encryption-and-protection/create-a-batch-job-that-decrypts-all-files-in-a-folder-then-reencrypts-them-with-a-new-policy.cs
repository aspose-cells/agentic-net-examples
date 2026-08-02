// Title: Batch Decrypt and Re‑Encrypt Excel Workbooks with a New Policy using Aspose.Cells for .NET
// Description: Scans a folder for .xlsx files, opens each workbook with the legacy password, applies a new password, encryption type, and key length, then saves the files to a target directory while logging any errors.
// Keywords: Aspose.Cells | C# batch encryption | Excel workbook decryption | change Excel password programmatically | set encryption type | strong cryptographic provider | bulk Excel encryption | file encryption policy | LoadOptions password | Workbook.SetEncryptionOptions
// Common Searches: C# batch decrypt Excel files Aspose.Cells | how to re‑encrypt multiple .xlsx files with new password | change encryption type of Excel workbooks programmatically | bulk update Excel file protection using Aspose.Cells | automate Excel encryption policy migration .NET
// Developer Intent: Decrypt every Excel file in a folder and re‑save it with a new password and encryption configuration.
// Use Cases: Upgrade legacy spreadsheets protected with weak passwords to a strong cryptographic provider across a corporate document library. | Run a scheduled job that enforces organization‑wide encryption standards on newly uploaded Excel reports. | Prepare a batch of workbooks for archiving by increasing key length to meet compliance requirements.
// AI Prompts: Create C# code that iterates over .xls and .xlsx files, removes existing protection, and applies AES‑256 encryption with a configurable password using Aspose.Cells. | Refactor the batch encryption script to accept input and output paths via command‑line arguments and generate a CSV log of processed files. | Write a PowerShell wrapper that calls the Aspose.Cells C# utility, passes folder parameters, and sends a summary email after completion.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Metadata;

// Scans a folder for .xlsx files, opens each workbook with the legacy password, applies a new password, encryption type, and key length, then saves the files to a target directory while logging any errors.
class BatchEncryptDecrypt
{
    static void Main()
    {
        // Folder containing the source Excel files
        string sourceFolder = @"C:\InputFiles";

        // Folder where the re‑encrypted files will be saved
        string outputFolder = @"C:\OutputFiles";

        // Old password used to open the encrypted files (if any)
        string oldPassword = "oldPassword";

        // New password and encryption policy to apply
        string newPassword = "newPassword";
        EncryptionType newEncryptionType = EncryptionType.StrongCryptographicProvider;
        int newKeyLength = 128; // 40, 128, or 256 bits are supported

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Verify source directory exists
        if (!Directory.Exists(sourceFolder))
        {
            Console.WriteLine($"Source folder not found: {sourceFolder}");
            return;
        }

        // Process each Excel file in the source folder
        foreach (string filePath in Directory.GetFiles(sourceFolder, "*.xlsx"))
        {
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found (skipped): {filePath}");
                continue;
            }

            try
            {
                // Load the workbook using the old password (if the file is encrypted)
                LoadOptions loadOptions = new LoadOptions
                {
                    Password = oldPassword
                };

                Workbook workbook = new Workbook(filePath, loadOptions);

                // Apply the new encryption settings
                workbook.Settings.Password = newPassword;
                workbook.SetEncryptionOptions(newEncryptionType, newKeyLength);

                // Save the workbook to the output folder (overwrites any existing file with the same name)
                string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
                workbook.Save(outputPath);
                Console.WriteLine($"Processed: {Path.GetFileName(filePath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("All files have been processed.");
    }
}
