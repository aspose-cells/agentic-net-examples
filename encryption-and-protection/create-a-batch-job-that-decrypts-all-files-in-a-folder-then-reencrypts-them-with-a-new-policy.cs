// Title: C# Batch Decrypt & Re‑encrypt Excel Files with Aspose.Cells – Update Password & Encryption Policy
// Description: A console utility that scans a source folder for .xlsx workbooks encrypted with an old password, opens each file using Aspose.Cells, applies a new password, switches to StrongCryptographicProvider with a 128‑bit key, and saves the re‑encrypted copies to a destination folder. Includes folder validation, per‑file error handling, and progress logging.
// Keywords: Aspose.Cells | C# batch encryption | Excel workbook decryption | re‑encrypt XLSX files | StrongCryptographicProvider | 128‑bit encryption key | change Excel password programmatically | process multiple Excel files | folder batch encryption | LoadOptions password
// Common Searches: how to batch decrypt encrypted Excel files using Aspose.Cells | C# code to change password of multiple XLSX workbooks | re‑encrypt Excel files with StrongCryptographicProvider in .NET | update encryption policy for a folder of Excel spreadsheets | automate Excel file password rotation with Aspose.Cells
// Developer Intent: Automate the decryption of all encrypted Excel workbooks in a directory and re‑save them with a new password and stronger encryption settings.
// Use Cases: Migrate legacy encrypted spreadsheets to a stronger encryption algorithm before long‑term storage. | Enforce organization‑wide password policies by rotating passwords on shared Excel reports. | Integrate into CI/CD pipelines to apply a consistent encryption policy to generated spreadsheets.
// AI Prompts: Generate C# code that iterates over a folder, opens each .xlsx with an existing password using Aspose.Cells, and saves it with a new password and StrongCryptographicProvider encryption. | Add robust logging and retry logic to the batch encryption script, capturing file‑level successes and failures. | Extend the utility to handle .xls and .xlsm formats and read encryption parameters from a JSON configuration file.

using System;
using System.IO;
using Aspose.Cells;

namespace BatchEncryption
{
    // A console utility that scans a source folder for .xlsx workbooks encrypted with an old password, opens each file using Aspose.Cells, applies a new password, switches to StrongCryptographicProvider with a 128‑bit key, and saves the re‑encrypted copies to a destination folder. Includes folder validation, per‑file error handling, and progress logging.
    class Program
    {
        static void Main()
        {
            // Folder containing the Excel files to process
            string sourceFolder = @"C:\InputFiles";
            // Folder where the re‑encrypted files will be saved
            string destinationFolder = @"C:\OutputFiles";

            // Verify source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder does not exist: {sourceFolder}");
                return;
            }

            // Ensure the destination folder exists
            Directory.CreateDirectory(destinationFolder);

            // Old password used for the existing encrypted files
            const string oldPassword = "oldPassword";
            // New password and encryption policy to apply
            const string newPassword = "newPassword";
            const EncryptionType newEncryptionType = EncryptionType.StrongCryptographicProvider;
            const int newKeyLength = 128; // 128‑bit key

            // Process each Excel file in the source folder
            foreach (string filePath in Directory.GetFiles(sourceFolder, "*.xlsx"))
            {
                try
                {
                    // Verify the file still exists before loading
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found: {filePath}");
                        continue;
                    }

                    // Load the workbook using the old password (decryption step)
                    LoadOptions loadOptions = new LoadOptions
                    {
                        Password = oldPassword
                    };
                    Workbook workbook = new Workbook(filePath, loadOptions);

                    // Apply the new encryption policy
                    workbook.Settings.Password = newPassword;
                    workbook.SetEncryptionOptions(newEncryptionType, newKeyLength);

                    // Build the output file path
                    string fileName = Path.GetFileName(filePath);
                    string outputPath = Path.Combine(destinationFolder, fileName);

                    // Save the workbook with the new encryption (re‑encryption step)
                    workbook.Save(outputPath, SaveFormat.Xlsx);

                    Console.WriteLine($"Processed: {fileName}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch encryption job completed.");
        }
    }
}
