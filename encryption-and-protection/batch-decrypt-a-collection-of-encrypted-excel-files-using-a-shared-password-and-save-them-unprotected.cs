// Title: Decrypt multiple password‑protected Excel files in a folder using Aspose.Cells for .NET
// Common Searches: C# code to decrypt all encrypted .xlsx files in a directory with Aspose.Cells | How to remove workbook password from a batch of Excel files using Aspose.Cells | Programmatically open password‑protected Excel workbooks and save them unprotected in .NET | Bulk decrypt Excel files with a shared password using Aspose.Cells LoadOptions | Decrypt a folder of protected Excel spreadsheets with a single password in C#
// Tags: batch workbook decryption Aspose.Cells | load encrypted Excel with password .NET | remove workbook password Aspose.Cells | save unprotected xlsx Aspose.Cells | process multiple Excel files C#

using System;
using System.IO;
using Aspose.Cells;

// The ExcelBatchDecryptor class iterates over .xls, .xlsx, and .xlsm files in a specified input folder, opens each workbook with a shared password via Aspose.Cells LoadOptions, clears any workbook‑level password, and saves the unprotected copy to an output folder.
class ExcelBatchDecryptor
{
    /// <param name="inputFolder">Folder containing encrypted Excel files.</param>
    /// <param name="outputFolder">Folder where decrypted files will be saved.</param>
    /// <param name="password">Shared password used to open the encrypted files.</param>
    public static void DecryptFolder(string inputFolder, string outputFolder, string password)
    {
        // Verify input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder does not exist: {inputFolder}");
            return;
        }

        // Ensure output directory exists
        if (!Directory.Exists(outputFolder))
            Directory.CreateDirectory(outputFolder);

        // Get all Excel files (xls, xlsx, xlsm) in the input folder
        string[] files = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
        foreach (string filePath in files)
        {
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension != ".xls" && extension != ".xlsx" && extension != ".xlsm")
                continue; // Skip non‑Excel files

            // Verify the file still exists before processing
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found (skipped): {filePath}");
                continue;
            }

            try
            {
                // Load the workbook with the provided password
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto)
                {
                    Password = password
                };
                Workbook workbook = new Workbook(filePath, loadOptions);

                // Remove any workbook-level password (if set)
                workbook.Settings.Password = null;

                // Build output file path (same name, different folder)
                string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));

                // Save the workbook without any password (unprotected)
                workbook.Save(outputPath);
                Console.WriteLine($"Decrypted: {Path.GetFileName(filePath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }
    }

    // Example usage
    static void Main()
    {
        string inputFolder = @"C:\EncryptedExcels";
        string outputFolder = @"C:\DecryptedExcels";
        string sharedPassword = "MySecretPassword";

        try
        {
            DecryptFolder(inputFolder, outputFolder, sharedPassword);
            Console.WriteLine("Decryption completed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
