// Title: C# Batch Decrypt Encrypted Excel (.xlsx) Files with a Shared Password Using Aspose.Cells
// Description: A complete C# console example that scans a source folder, loads each password‑protected .xlsx workbook with Aspose.Cells LoadOptions, removes workbook protection, clears the encryption password, and saves the unprotected copy to a destination folder. Includes folder validation, robust error handling, and progress logging.
// Keywords: Aspose.Cells | C# batch decrypt Excel | remove Excel password programmatically | load encrypted workbook Aspose | unprotect multiple .xlsx files | Excel encryption removal .NET | bulk Excel decryption | shared password Excel | console app Aspose.Cells | folder processing C#
// Common Searches: batch decrypt Excel files C# | remove password from multiple .xlsx using Aspose.Cells | C# code to unprotect encrypted workbooks in a folder | how to bulk decrypt Excel workbooks .NET | Aspose.Cells load encrypted workbook with password | automate Excel password removal C#
// Developer Intent: Programmatically open a set of password‑protected Excel workbooks, strip their protection, and save them unencrypted.
// Use Cases: Automated nightly decryption of secured financial reports before ETL processing. | Pre‑processing client‑submitted spreadsheets for import into ERP systems that reject protected files. | Bulk migration of legacy encrypted Excel archives to a plain‑text repository for compliance audits. | Generating unprotected copies for data‑science pipelines that require direct cell access.
// AI Prompts: Write a C# console program that iterates over all .xlsx files in a given directory, opens each with a shared password using Aspose.Cells, removes workbook protection, clears the password, and saves the result to another directory, with robust error handling. | Show how to log each successful decryption and capture exceptions when a file cannot be opened due to an incorrect password, using Aspose.Cells in .NET. | Create a PowerShell wrapper that calls the compiled C# batch decryption tool for scheduled tasks. | Explain how to modify the example to support different passwords per file using a CSV mapping.

using System;
using System.IO;
using Aspose.Cells;

namespace BatchDecryptExcel
{
    // A complete C# console example that scans a source folder, loads each password‑protected .xlsx workbook with Aspose.Cells LoadOptions, removes workbook protection, clears the encryption password, and saves the unprotected copy to a destination folder. Includes folder validation, robust error handling, and progress logging.
    class Program
    {
        static void Main()
        {
            // Folder containing encrypted Excel files
            string inputFolder = @"C:\EncryptedFiles";
            // Folder where unprotected files will be saved
            string outputFolder = @"C:\DecryptedFiles";

            // Verify input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Shared password for all encrypted workbooks
            string sharedPassword = "MySharedPassword";

            // Process each .xlsx file in the input folder
            foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
            {
                try
                {
                    // Verify the file still exists
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found, skipping: {filePath}");
                        continue;
                    }

                    // Load the workbook with the password
                    LoadOptions loadOptions = new LoadOptions
                    {
                        Password = sharedPassword
                    };
                    Workbook workbook = new Workbook(filePath, loadOptions);

                    // Remove workbook protection (if any)
                    workbook.Unprotect(sharedPassword);

                    // Ensure the workbook is no longer encrypted when saved
                    workbook.Settings.Password = null;

                    // Build the output file path
                    string fileName = Path.GetFileName(filePath);
                    string outputPath = Path.Combine(outputFolder, fileName);

                    // Save the unprotected workbook
                    workbook.Save(outputPath);

                    Console.WriteLine($"Decrypted: {fileName}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch decryption completed.");
        }
    }
}
