// Title: C# Batch Encrypt Excel Files with Aspose.Cells – Password Derived from File Name
// Description: A C# console utility that scans a folder for Excel workbooks (*.xlsx, *.xls, *.xlsm, *.xlsb), loads each file with Aspose.Cells, generates a password from the file name (without extension), applies the password (and 128‑bit strong encryption for legacy .xls files), and saves the encrypted workbook back to the original location.
// Keywords: Aspose.Cells batch encryption | C# encrypt multiple Excel files | password from file name | Excel workbook protection .NET | strong encryption for .xls | automated Excel security | folder scan Excel encryption | Aspose.Cells password protection example | C# console app encrypt Excel | bulk Excel file encryption
// Common Searches: how to encrypt all Excel files in a folder using Aspose.Cells | C# set workbook password programmatically | batch encrypt .xls files with 128‑bit encryption | Aspose.Cells encrypt multiple workbooks | derive Excel password from file name C#
// Developer Intent: Automatically protect every Excel workbook in a directory by assigning a unique password based on its file name.
// Use Cases: Secure a collection of financial spreadsheets before archiving, giving each file a distinct password derived from its name. | Run nightly compliance jobs that encrypt legacy .xls reports with strong 128‑bit encryption. | Prepare Excel deliverables for external partners, ensuring each recipient can open only their file using the filename as the password.
// AI Prompts: Write a C# program using Aspose.Cells that encrypts all Excel files in a specified folder, using the file name (without extension) as the password and applying 128‑bit encryption for .xls files. | Refactor the batch encryption code to log successes and failures to a CSV file and skip files that already have a password set. | Explain how to replace the plain‑text filename password with a SHA‑256 hash generated from the file name.

using System;
using System.IO;
using Aspose.Cells;

namespace ExcelBatchEncryption
{
    // A C# console utility that scans a folder for Excel workbooks (*.xlsx, *.xls, *.xlsm, *.xlsb), loads each file with Aspose.Cells, generates a password from the file name (without extension), applies the password (and 128‑bit strong encryption for legacy .xls files), and saves the encrypted workbook back to the original location.
    class BatchEncryptExcel
    {
        static void Main()
        {
            // Folder containing the Excel files to encrypt
            string folderPath = @"C:\ExcelFolder";

            // Verify that the folder exists
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder not found: {folderPath}");
                return;
            }

            // Supported Excel file extensions
            string[] extensions = new[] { "*.xlsx", "*.xls", "*.xlsm", "*.xlsb" };

            foreach (string ext in extensions)
            {
                string[] files;
                try
                {
                    files = Directory.GetFiles(folderPath, ext);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error retrieving files with pattern '{ext}': {ex.Message}");
                    continue;
                }

                foreach (string filePath in files)
                {
                    // Ensure the file still exists before processing
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found: {filePath}");
                        continue;
                    }

                    try
                    {
                        // Load workbook inside a using block for automatic disposal
                        using (Workbook workbook = new Workbook(filePath))
                        {
                            // Derive a password from the file name (without extension)
                            string password = Path.GetFileNameWithoutExtension(filePath);

                            // Apply password protection (encryption)
                            workbook.Settings.Password = password;

                            // For legacy .xls files, set stronger encryption options
                            if (Path.GetExtension(filePath).Equals(".xls", StringComparison.OrdinalIgnoreCase))
                            {
                                workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);
                            }

                            // Overwrite the original file with the encrypted version
                            workbook.Save(filePath);
                        }

                        Console.WriteLine($"Encrypted '{Path.GetFileName(filePath)}' with password '{Path.GetFileNameWithoutExtension(filePath)}'.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to encrypt '{filePath}': {ex.Message}");
                    }
                }
            }
        }
    }
}
