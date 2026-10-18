// Title: How to batch encrypt XLS workbooks with a shared password using Aspose.Cells for .NET
// Common Searches: C# code to encrypt all .xls files in a folder with Aspose.Cells and save to another directory | Aspose.Cells batch password protection for Excel 97-2003 files in .NET | Programmatically apply the same password to multiple XLS workbooks using Aspose.Cells | Encrypt multiple legacy Excel files (.xls) in bulk with a common password in C# | Save encrypted XLS files to a different folder using Aspose.Cells API
// Tags: Aspose.Cells batch XLS encryption | C# set workbook password Aspose.Cells | encrypt Excel97To2003 files programmatically | bulk protect legacy Excel workbooks .NET | save encrypted workbooks to target folder

using System;
using System.IO;
using Aspose.Cells;

// The example scans a source directory for .xls files, loads each workbook with Aspose.Cells, assigns a common password via Workbook.Settings.Password, and saves the encrypted workbook to a target folder in Excel97To2003 format.
class BatchEncryptXls
{
    static void Main()
    {
        // Source folder containing the original XLS files
        string sourceFolder = @"C:\SourceFolder";

        // Target folder where encrypted files will be saved
        string targetFolder = @"C:\TargetFolder";

        // Common password to encrypt all workbooks
        string password = "CommonPassword";

        try
        {
            // Verify source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder not found: {sourceFolder}");
                return;
            }

            // Ensure the target directory exists
            Directory.CreateDirectory(targetFolder);

            // Retrieve all .xls files from the source folder (non‑recursive)
            string[] xlsFiles = Directory.GetFiles(sourceFolder, "*.xls", SearchOption.TopDirectoryOnly);

            foreach (string sourcePath in xlsFiles)
            {
                // Skip if the file somehow does not exist
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"File not found, skipping: {sourcePath}");
                    continue;
                }

                try
                {
                    // Load the workbook from the source file
                    Workbook workbook = new Workbook(sourcePath);

                    // Apply password protection (encryption)
                    workbook.Settings.Password = password;

                    // Determine the target file path (same file name, different folder)
                    string fileName = Path.GetFileName(sourcePath);
                    string targetPath = Path.Combine(targetFolder, fileName);

                    // Save the encrypted workbook to the target location in XLS format
                    workbook.Save(targetPath, SaveFormat.Excel97To2003);
                    Console.WriteLine($"Encrypted and saved: {targetPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{sourcePath}': {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
