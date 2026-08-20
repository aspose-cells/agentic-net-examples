// Title: Batch encrypt Excel .xlsx files with Aspose.Cells for .NET using filename‑derived passwords
// Description: C# example that scans a directory, loads each .xlsx workbook via Aspose.Cells, sets the workbook password to the file name (without extension), applies AES‑128 strong encryption, and saves the file back, overwriting the original. Includes folder validation and exception handling.
// Keywords: Aspose.Cells encrypt Excel | C# batch Excel encryption | filename password Excel | AES 128 Aspose.Cells | protect .xlsx files programmatically | Excel workbook password .NET | bulk Excel security
// Common Searches: how to encrypt multiple xlsx files with Aspose.Cells | set Excel workbook password from file name using C# | batch apply AES encryption to Excel files | Aspose.Cells bulk password protection example | automate Excel file encryption .NET
// Developer Intent: Programmatically encrypt every Excel workbook in a folder, assigning each a unique password derived from its file name.
// Use Cases: Secure a batch of financial reports before archiving them on a shared drive. | Automatically protect generated invoices so each can be opened only with its own identifier as the password. | Encrypt confidential spreadsheets prior to uploading them to cloud storage or a collaboration platform.
// AI Prompts: Create C# code with Aspose.Cells that encrypts all .xls files in a directory, using a password pattern of "Report_" + file name, and logs any failures. | Show how to switch the encryption to AES‑256 and save the encrypted workbooks to a separate output folder. | Explain how to modify the sample to skip files that are already password‑protected and to generate a summary report after processing.

using System;
using System.IO;
using Aspose.Cells;

namespace BatchEncryptExcel
{
    // C# example that scans a directory, loads each .xlsx workbook via Aspose.Cells, sets the workbook password to the file name (without extension), applies AES‑128 strong encryption, and saves the file back, overwriting the original. Includes folder validation and exception handling.
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

            // Get all Excel files in the folder (adjust the pattern as needed)
            string[] excelFiles = Directory.GetFiles(folderPath, "*.xlsx");

            foreach (string filePath in excelFiles)
            {
                // Ensure the file still exists before processing
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook from the existing file
                    using (Workbook workbook = new Workbook(filePath))
                    {
                        // Derive a unique password from the file name (without extension)
                        string password = Path.GetFileNameWithoutExtension(filePath);

                        // Set the password for the workbook
                        workbook.Settings.Password = password;

                        // Optional: set stronger encryption options (AES 128-bit)
                        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

                        // Save the workbook back to the same file (overwrites the original)
                        workbook.Save(filePath, SaveFormat.Xlsx);
                    }

                    Console.WriteLine($"Encrypted '{Path.GetFileName(filePath)}' with password '{Path.GetFileNameWithoutExtension(filePath)}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{Path.GetFileName(filePath)}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch encryption completed.");
        }
    }
}
