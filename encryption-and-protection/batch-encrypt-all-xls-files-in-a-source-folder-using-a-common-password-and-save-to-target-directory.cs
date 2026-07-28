// Title: C# Example: Batch Encrypt XLS Files with a Shared Password Using Aspose.Cells
// Description: This C# .NET console example scans a source folder for *.xls workbooks, loads each file with Aspose.Cells, applies a common password, enforces strong 128‑bit encryption, and saves the protected copies to a target directory. It creates missing folders, skips unavailable files, and logs progress to the console.
// Keywords: Aspose.Cells | C# batch encrypt XLS | Excel password protection .NET | encrypt multiple Excel files | strong encryption Aspose | set workbook password | save encrypted workbook | Windows file processing | console application | legacy .xls encryption
// Common Searches: batch encrypt xls files c# | aspocells encrypt multiple workbooks | set password for all Excel files programmatically | c# encrypt legacy .xls with Aspose.Cells | strong encryption for Excel files .NET | how to encrypt Excel files in a folder using Aspose | encrypt xls files and save to another folder c#
// Developer Intent: Apply a single password to every .xls workbook in a folder and write the encrypted files to a separate directory.
// Use Cases: Protect a collection of legacy reports before archiving to a secure share. | Add company‑wide password protection to automatically generated spreadsheets in a nightly batch job. | Migrate unencrypted workbooks to a compliance‑required encrypted repository.
// AI Prompts: Write C# code that uses Aspose.Cells to encrypt all .xls files in a specified folder with a given password and output them to another folder. | Modify the example to use AES‑256 encryption and include a per‑file random salt. | Add functionality to log each file’s encryption result (success or error) to a CSV file.

using System;
using System.IO;
using Aspose.Cells;

namespace BatchEncryptXls
{
    // This C# .NET console example scans a source folder for *.xls workbooks, loads each file with Aspose.Cells, applies a common password, enforces strong 128‑bit encryption, and saves the protected copies to a target directory. It creates missing folders, skips unavailable files, and logs progress to the console.
    class Program
    {
        static void Main(string[] args)
        {
            // Define source folder, target folder and the common password
            string sourceFolder = @"C:\SourceFolder";
            string targetFolder = @"C:\TargetFolder";
            string password = "CommonPassword123";

            try
            {
                // Verify source folder exists
                if (!Directory.Exists(sourceFolder))
                {
                    Console.WriteLine($"Source folder does not exist: {sourceFolder}");
                    return;
                }

                // Ensure target folder exists
                if (!Directory.Exists(targetFolder))
                {
                    Directory.CreateDirectory(targetFolder);
                }

                // Get all .xls files in the source folder
                string[] xlsFiles = Directory.GetFiles(sourceFolder, "*.xls", SearchOption.TopDirectoryOnly);

                foreach (string sourceFilePath in xlsFiles)
                {
                    try
                    {
                        // Verify the source file still exists
                        if (!File.Exists(sourceFilePath))
                        {
                            Console.WriteLine($"File not found, skipping: {sourceFilePath}");
                            continue;
                        }

                        // Load the workbook from the source file
                        Workbook workbook = new Workbook(sourceFilePath);

                        // Set the password for the workbook (encryption)
                        workbook.Settings.Password = password;

                        // Optional: specify encryption algorithm and key length (strong encryption)
                        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

                        // Determine the destination file path
                        string fileName = Path.GetFileName(sourceFilePath);
                        string destFilePath = Path.Combine(targetFolder, fileName);

                        // Save the encrypted workbook to the target directory
                        workbook.Save(destFilePath);
                        Console.WriteLine($"Encrypted: {fileName}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{sourceFilePath}': {ex.Message}");
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
