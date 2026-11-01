// Title: Encrypt all Excel workbooks in a SharePoint library with a shared password using Aspose.Cells for .NET
// Common Searches: C# code to apply the same opening password to every Excel file in a SharePoint folder with Aspose.Cells | how to batch encrypt .xlsx and .xls files from a SharePoint document library using Aspose.Cells .NET | programmatically protect multiple workbooks with a central password in a SharePoint library using Aspose.Cells | Aspose.Cells encrypt Excel files while preserving folder hierarchy in C# | set workbook.Settings.Password for bulk Excel files from SharePoint using .NET
// Tags: Aspose.Cells batch workbook encryption | shared opening password for Excel files | encrypt Excel files from SharePoint library | preserve directory structure when saving encrypted workbooks | C# Aspose.Cells SaveFormat.Xlsx encryption

using System;
using System.IO;
using Aspose.Cells;

namespace SharePointWorkbookEncryption
{
    // The example scans a SharePoint‑mapped folder, loads each .xls, .xlsx, or .xlsm workbook with Aspose.Cells, assigns a common opening password via workbook.Settings.Password, and saves the encrypted files to an output location while keeping the original folder hierarchy intact.
    class Program
    {
        // Centralized password for all workbooks
        private const string WorkbookPassword = "YourSecurePassword";

        // Local folders used for demonstration (replace with actual paths)
        private const string InputFolder = @"C:\InputExcel";
        private const string OutputFolder = @"C:\EncryptedExcel";

        static void Main()
        {
            try
            {
                // Ensure input folder exists
                if (!Directory.Exists(InputFolder))
                {
                    Console.WriteLine($"Input folder does not exist: {InputFolder}");
                    return;
                }

                // Create output folder if it does not exist
                if (!Directory.Exists(OutputFolder))
                {
                    Directory.CreateDirectory(OutputFolder);
                }

                // Get all Excel files in the input folder (including subfolders)
                string[] excelFiles = Directory.GetFiles(InputFolder, "*.*", SearchOption.AllDirectories);
                foreach (string filePath in excelFiles)
                {
                    string extension = Path.GetExtension(filePath);
                    if (!extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) &&
                        !extension.Equals(".xls", StringComparison.OrdinalIgnoreCase) &&
                        !extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase))
                    {
                        continue; // Skip non‑Excel files
                    }

                    string fileName = Path.GetFileName(filePath);
                    Console.WriteLine($"Encrypting: {fileName}");

                    try
                    {
                        // Verify the file exists before attempting to open it
                        if (!File.Exists(filePath))
                        {
                            Console.WriteLine($"File not found: {filePath}");
                            continue;
                        }

                        // Load workbook from file
                        using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                        {
                            Workbook workbook = new Workbook(fs);

                            // Set the opening password (this encrypts the file)
                            workbook.Settings.Password = WorkbookPassword;

                            // Determine output path preserving folder structure
                            string relativePath = Path.GetRelativePath(InputFolder, filePath);
                            string outputPath = Path.Combine(OutputFolder, relativePath);
                            string outputDir = Path.GetDirectoryName(outputPath);
                            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                            {
                                Directory.CreateDirectory(outputDir);
                            }

                            // Save the encrypted workbook using the appropriate format
                            if (extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                                extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase))
                            {
                                workbook.Save(outputPath, SaveFormat.Xlsx);
                            }
                            else // .xls
                            {
                                workbook.Save(outputPath, SaveFormat.Excel97To2003);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{fileName}': {ex.Message}");
                    }
                }

                Console.WriteLine("Encryption process completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
