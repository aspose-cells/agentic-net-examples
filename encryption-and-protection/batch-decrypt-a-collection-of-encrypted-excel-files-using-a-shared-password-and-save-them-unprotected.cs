// Title: Batch decrypt multiple encrypted .xlsx workbooks using a shared password with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loops through every .xlsx file in a directory, opens each workbook with LoadOptions.Password set to a common password, calls Workbook.Unprotect, and saves the result without a password to a target folder using Aspose.Cells. | Write a .NET console application that reads password‑protected Excel files from a source folder, removes the workbook protection using the same password, and writes unprotected copies to an output folder with Aspose.Cells.
// Common Searches: aspnet batch remove password from Excel files using Aspose.Cells | C# program to decrypt multiple encrypted .xlsx files with the same password | how to use LoadOptions.Password to open encrypted workbooks in a loop | save unprotected copy of encrypted Excel workbook with Aspose.Cells .NET | process folder of password‑protected Excel files in C#
// Tags: decrypt multiple xlsx files Aspose.Cells | load encrypted workbook with password .NET | remove workbook protection programmatically | save unprotected Excel workbook C# | iterate Excel files in folder Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The C# console program scans a given input directory for .xlsx files, opens each encrypted workbook using a shared password via Aspose.Cells LoadOptions, removes workbook protection, and saves the unprotected files to a specified output directory.
class Program
{
    static void Main()
    {
        // Folder containing the encrypted Excel files
        string inputFolder = @"C:\EncryptedExcels";

        // Folder where the decrypted (unprotected) files will be saved
        string outputFolder = @"C:\DecryptedExcels";

        // Shared password used to open and unprotect all encrypted workbooks
        string password = "sharedPassword";

        try
        {
            // Verify input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Ensure output folder exists
            Directory.CreateDirectory(outputFolder);

            // Retrieve all .xlsx files in the input folder
            string[] encryptedFiles = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string encryptedFilePath in encryptedFiles)
            {
                try
                {
                    // Confirm the file still exists before attempting to load
                    if (!File.Exists(encryptedFilePath))
                    {
                        Console.WriteLine($"File not found: {encryptedFilePath}");
                        continue;
                    }

                    // Load the workbook with the shared password
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                    {
                        Password = password
                    };
                    Workbook workbook = new Workbook(encryptedFilePath, loadOptions);

                    // Remove workbook protection using the same password
                    workbook.Unprotect(password);

                    // Save the workbook without a password
                    string fileName = Path.GetFileName(encryptedFilePath);
                    string decryptedFilePath = Path.Combine(outputFolder, fileName);
                    workbook.Save(decryptedFilePath, SaveFormat.Xlsx);

                    Console.WriteLine($"Decrypted and saved: {decryptedFilePath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{encryptedFilePath}': {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
