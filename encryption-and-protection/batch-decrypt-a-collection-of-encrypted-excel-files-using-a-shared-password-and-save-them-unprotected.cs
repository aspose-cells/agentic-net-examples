// Title: Batch Decrypt Encrypted Excel Files with a Shared Password Using Aspose.Cells for .NET (C#)
// Description: C# code that scans a folder, opens each password‑protected XLSX, XLS, XLSM or XLSB workbook with a common password via Aspose.Cells LoadOptions, removes the encryption, and saves the unprotected copy to a target directory. Includes folder validation and per‑file error handling.
// Keywords: Aspose.Cells batch decrypt | C# remove Excel password | decrypt multiple XLSX files | load encrypted workbook Aspose.Cells | save unprotected Excel .NET | bulk Excel decryption C# | shared password Excel decryption
// Common Searches: how to batch decrypt password protected Excel files using Aspose.Cells | c# code to remove encryption from many XLSX workbooks | bulk Excel decryption with a single password .NET | Aspose.Cells example for decrypting multiple spreadsheets | programmatically open encrypted Excel files in C#
// Developer Intent: Remove password protection from a collection of Excel workbooks in one operation and write the decrypted files to a new folder.
// Use Cases: Automate the cleanup of archived reports that were saved with a corporate password before publishing. | Prepare encrypted financial models for data‑analysis pipelines that require unprotected files. | Migrate legacy workbooks to a new system where password protection is no longer needed.
// AI Prompts: Write C# code with Aspose.Cells to batch decrypt Excel files in a directory using a shared password. | Explain how to log and skip files that fail decryption when processing many workbooks. | Show how to preserve original file timestamps after saving the decrypted Excel files.

using System;
using System.IO;
using Aspose.Cells;

namespace BatchDecryptExample
{
    // C# code that scans a folder, opens each password‑protected XLSX, XLS, XLSM or XLSB workbook with a common password via Aspose.Cells LoadOptions, removes the encryption, and saves the unprotected copy to a target directory. Includes folder validation and per‑file error handling.
    public class Decryptor
    {
        /// <param name="inputFolder">Folder containing encrypted Excel files.</param>
        /// <param name="outputFolder">Folder where unprotected files will be saved.</param>
        /// <param name="password">Shared password used to open the encrypted files.</param>
        public static void Run(string inputFolder, string outputFolder, string password)
        {
            try
            {
                // Verify input folder exists
                if (!Directory.Exists(inputFolder))
                {
                    Console.WriteLine($"Input folder does not exist: {inputFolder}");
                    return;
                }

                // Ensure the output directory exists
                if (!Directory.Exists(outputFolder))
                {
                    Directory.CreateDirectory(outputFolder);
                }

                // Get all Excel files (XLSX, XLS, XLSM, XLSB) in the input folder
                string[] excelFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
                foreach (string filePath in excelFiles)
                {
                    // Filter supported Excel extensions
                    string ext = Path.GetExtension(filePath).ToLowerInvariant();
                    if (ext != ".xlsx" && ext != ".xls" && ext != ".xlsm" && ext != ".xlsb")
                    {
                        continue;
                    }

                    // Ensure the file actually exists before processing
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found: {filePath}");
                        continue;
                    }

                    try
                    {
                        // Load the encrypted workbook using the shared password
                        LoadOptions loadOptions = new LoadOptions
                        {
                            Password = password // set password for encrypted file
                        };
                        Workbook workbook = new Workbook(filePath, loadOptions);

                        // Remove any encryption password before saving
                        workbook.Settings.Password = null;

                        // Determine the output file path
                        string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));

                        // Save the unprotected workbook
                        workbook.Save(outputPath);
                        Console.WriteLine($"Decrypted and saved: {outputPath}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to process '{filePath}': {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error during batch decryption: {ex.Message}");
            }
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            string inputFolder = @"C:\EncryptedExcels";
            string outputFolder = @"C:\DecryptedExcels";
            string sharedPassword = "MySharedPassword";

            Decryptor.Run(inputFolder, outputFolder, sharedPassword);
            Console.WriteLine("Batch decryption completed.");
        }
    }
}
