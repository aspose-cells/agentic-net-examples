// Title: Batch encrypt multiple XLS (Excel 97‑2003) workbooks with a shared password using Aspose.Cells for .NET
// Common Searches: C# code to encrypt all .xls files in a folder with Aspose.Cells | How to apply the same password to many Excel 97-2003 workbooks using Aspose.Cells .NET | Batch protect Excel files with a password programmatically in C# | Encrypt a directory of .xls files and save them to another location using Aspose.Cells | Aspose.Cells example for bulk workbook password protection in .NET
// Tags: batch encrypt xls files Aspose.Cells | set workbook password Aspose.Cells C# | encrypt multiple Excel 97-2003 workbooks .NET | save encrypted workbooks to target folder C# | apply common password to Excel files Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample scans a source directory for .xls files, loads each workbook with Aspose.Cells, assigns a common password via Workbook.Settings.Password, and saves the encrypted copies to a target folder, creating the destination if needed and handling file‑level and unexpected errors.
class Program
{
    static void Main()
    {
        // Source folder containing original XLS files
        string sourceFolder = @"C:\SourceFolder";

        // Target folder where encrypted files will be saved
        string targetFolder = @"C:\TargetFolder";

        // Common password for all encrypted workbooks
        string password = "MyPassword";

        try
        {
            // Ensure the source directory exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder does not exist: {sourceFolder}");
                return;
            }

            // Ensure the target directory exists
            Directory.CreateDirectory(targetFolder);

            // Retrieve all .xls files from the source folder (non‑recursive)
            string[] xlsFiles = Directory.GetFiles(sourceFolder, "*.xls", SearchOption.TopDirectoryOnly);

            foreach (string sourcePath in xlsFiles)
            {
                try
                {
                    // Verify the source file exists before loading
                    if (!File.Exists(sourcePath))
                    {
                        Console.WriteLine($"File not found: {sourcePath}");
                        continue;
                    }

                    // Load the workbook from the source file
                    Workbook workbook = new Workbook(sourcePath);

                    // Apply password protection to the workbook
                    workbook.Settings.Password = password;

                    // Determine the target file path (same file name, different folder)
                    string fileName = Path.GetFileName(sourcePath);
                    string targetPath = Path.Combine(targetFolder, fileName);

                    // Save the encrypted workbook to the target location
                    workbook.Save(targetPath);
                    Console.WriteLine($"Encrypted file saved: {targetPath}");
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
