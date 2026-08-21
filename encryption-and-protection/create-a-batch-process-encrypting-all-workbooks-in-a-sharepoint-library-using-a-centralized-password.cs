// Title: Batch encrypt Excel and ODS workbooks in a SharePoint library with Aspose.Cells for .NET
// Description: C# utility that scans a SharePoint (or local) document library, loads each workbook (xls, xlsx, xlsm, xlsb, ods) via Aspose.Cells, applies a single strong password, and saves the file back in its original format. Includes recursive folder search and robust error handling.
// Keywords: Aspose.Cells | C# | .NET | batch workbook encryption | Excel password protection | SharePoint document library | StrongCryptographicProvider | 128‑bit encryption | ODS encryption | automated compliance | file‑level security
// Common Searches: C# batch encrypt Excel files in SharePoint | Aspose.Cells set same password for multiple workbooks | How to encrypt all spreadsheets in a folder using .NET | Encrypt SharePoint library Excel and ODS files programmatically | Apply 128‑bit encryption to Excel workbooks with Aspose.Cells
// Developer Intent: Programmatically protect every workbook in a SharePoint document library with a unified password using Aspose.Cells for .NET.
// Use Cases: Secure confidential financial reports stored in SharePoint before distribution. | Run a nightly job that automatically password‑protects all newly uploaded spreadsheets to meet regulatory compliance. | Maintain original file formats while enforcing consistent encryption strength across mixed Excel and ODS files.
// AI Prompts: Generate C# code that iterates through a SharePoint document library, loads each workbook with Aspose.Cells, sets a common password, and saves it preserving the original format. | Show an Aspose.Cells for .NET example that encrypts a collection of Excel and ODS files using StrongCryptographicProvider with 128‑bit encryption and includes detailed error handling. | Explain how to extend the batch encryption script to skip already encrypted workbooks and write a processing log to a CSV file.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Cells;

namespace SharePointWorkbookEncryption
{
    // C# utility that scans a SharePoint (or local) document library, loads each workbook (xls, xlsx, xlsm, xlsb, ods) via Aspose.Cells, applies a single strong password, and saves the file back in its original format. Includes recursive folder search and robust error handling.
    class Program
    {
        // Centralized password for all workbooks
        private const string WorkbookPassword = "CentralPassword123";

        // Local folder that simulates the SharePoint library (replace with actual path)
        private const string LocalFolderPath = @"C:\Temp\Workbooks";

        static void Main()
        {
            try
            {
                // Verify the folder exists
                if (!Directory.Exists(LocalFolderPath))
                {
                    Console.WriteLine($"Folder not found: {LocalFolderPath}");
                    return;
                }

                // Get all workbook files in the folder and subfolders
                IEnumerable<string> files = Directory.EnumerateFiles(
                    LocalFolderPath,
                    "*.*",
                    SearchOption.AllDirectories)
                    .Where(f => new[] { ".xls", ".xlsx", ".xlsm", ".xlsb", ".ods" }
                    .Contains(Path.GetExtension(f).ToLowerInvariant()));

                foreach (string filePath in files)
                {
                    try
                    {
                        Console.WriteLine($"Encrypting workbook: {Path.GetFileName(filePath)}");

                        // Ensure the file exists before loading
                        if (!File.Exists(filePath))
                        {
                            Console.WriteLine($"File not found, skipping: {filePath}");
                            continue;
                        }

                        // Load the workbook (no password needed for loading unencrypted files)
                        using (FileStream inputStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                        {
                            Workbook workbook = new Workbook(inputStream);

                            // Set the encryption password
                            workbook.Settings.Password = WorkbookPassword;

                            // Optional: define encryption strength (ignored for modern formats but kept for completeness)
                            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

                            // Determine original format based on extension
                            string extension = Path.GetExtension(filePath).ToLowerInvariant();
                            SaveFormat format = GetSaveFormat(extension);

                            // Save the encrypted workbook back to the same file
                            using (FileStream outputStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                            {
                                workbook.Save(outputStream, format);
                            }

                            workbook.Dispose();
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                    }
                }

                Console.WriteLine("Encryption process completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        // Helper to map file extension to Aspose.Cells SaveFormat
        private static SaveFormat GetSaveFormat(string extension)
        {
            return extension switch
            {
                ".xls" => SaveFormat.Excel97To2003,
                ".xlsx" => SaveFormat.Xlsx,
                ".xlsm" => SaveFormat.Xlsm,
                ".xlsb" => SaveFormat.Xlsb,
                ".ods" => SaveFormat.Ods,
                _ => SaveFormat.Xlsx,
            };
        }
    }
}
