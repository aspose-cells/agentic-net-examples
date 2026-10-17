// Title: Decrypt multiple password‑protected Excel .xlsx files in batch with Aspose.Cells for .NET
// Common Searches: C# code to decrypt a list of password protected Excel files using Aspose.Cells | How to programmatically remove passwords from several .xlsx workbooks with Aspose.Cells .NET | Batch processing encrypted Excel files with different passwords in Aspose.Cells | Load encrypted workbook with password and save unprotected using Aspose.Cells in C# | Automate decryption of multiple Excel files stored in a folder using Aspose.Cells
// Tags: Aspose.Cells bulk Excel decryption | C# load password protected workbook Aspose.Cells | Aspose.Cells save workbook without password | dictionary driven Excel password mapping C# | automated removal of Excel file protection Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace ExcelDecryptionBatch
{
    // The example shows how to decrypt a batch of password‑protected Excel .xlsx files using Aspose.Cells for .NET. It stores each encrypted file path together with its password in a C# dictionary, creates an output folder, and iterates through the entries. For each file a LoadOptions object supplies the password, the workbook is loaded, and then saved to the output directory without a password, effectively removing protection. Errors are caught and logged, and a completion message is displayed.
    class Program
    {
        static void Main(string[] args)
        {
            // Dictionary mapping encrypted file paths to their passwords
            var filesWithPasswords = new Dictionary<string, string>
            {
                // Example entries – replace with actual paths and passwords
                { @"C:\EncryptedFiles\Report1.xlsx", "Password123" },
                { @"C:\EncryptedFiles\Report2.xlsx", "Secret!@#" },
                // Add more files as needed
            };

            // Folder where decrypted files will be saved
            string outputFolder = @"C:\DecryptedFiles";
            Directory.CreateDirectory(outputFolder);

            foreach (var kvp in filesWithPasswords)
            {
                string encryptedPath = kvp.Key;
                string password = kvp.Value;

                try
                {
                    // Load the encrypted workbook using the provided password
                    var loadOptions = new LoadOptions(LoadFormat.Xlsx)
                    {
                        Password = password
                    };
                    var workbook = new Workbook(encryptedPath, loadOptions);

                    // Determine output file name (same name, different folder)
                    string fileName = Path.GetFileName(encryptedPath);
                    string decryptedPath = Path.Combine(outputFolder, fileName);

                    // Save the workbook without a password (decrypted)
                    workbook.Save(decryptedPath, SaveFormat.Xlsx);

                    Console.WriteLine($"Decrypted and saved: {decryptedPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to decrypt '{encryptedPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch decryption completed.");
        }
    }
}
