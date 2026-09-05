// Title: Batch encrypt XLS (Excel 97‑2003) workbooks with a single password using Aspose.Cells for .NET
// AI Prompts: Write a C# program that scans a given folder for *.xls files, sets the same Workbook.Settings.Password for each workbook, and saves the encrypted files to a separate output directory in Excel97To2003 format. | Create code that ensures the target folder exists, loads each XLS workbook with Aspose.Cells, applies a common password, and writes the protected file while preserving the original file name. | Add logging to the batch routine so it records the name of each processed file and continues processing even if an individual file fails.
// Common Searches: asp.net batch encrypt xls files with Aspose.Cells | c# set password for multiple Excel 97-2003 workbooks | how to protect all .xls files in a folder using Aspose.Cells library | programmatically apply same password to many XLS files in C#
// Tags: batch encrypt xls Aspose.Cells | set workbook password Excel97To2003 C# | save encrypted xls files .NET | process multiple excel files directory Aspose | apply common password to Excel 97-2003 workbooks

using System;
using System.IO;
using Aspose.Cells;

// The example iterates over every .xls file in a source directory, loads each workbook with Aspose.Cells, assigns a shared password via Workbook.Settings.Password, and saves the protected workbook to a target folder using the Excel97To2003 format, creating the output folder if necessary and handling errors per file.
class Program
{
    static void Main()
    {
        try
        {
            // Source folder containing the XLS files
            string sourceFolder = @"C:\SourceFolder";
            // Target folder where encrypted files will be saved
            string targetFolder = @"C:\TargetFolder";
            // Common password for encryption
            string password = "MySecretPassword";

            // Ensure the target directory exists
            if (!Directory.Exists(targetFolder))
            {
                Directory.CreateDirectory(targetFolder);
            }

            // Verify source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder does not exist: {sourceFolder}");
                return;
            }

            // Get all .xls files from the source folder
            string[] xlsFiles = Directory.GetFiles(sourceFolder, "*.xls", SearchOption.TopDirectoryOnly);

            foreach (string filePath in xlsFiles)
            {
                try
                {
                    // Ensure the source file exists
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found: {filePath}");
                        continue;
                    }

                    // Load the workbook from the source file
                    Workbook workbook = new Workbook(filePath);

                    // Set password for the workbook (applies to XLS format)
                    workbook.Settings.Password = password;

                    // Determine the target file path (same file name, different folder)
                    string targetPath = Path.Combine(targetFolder, Path.GetFileName(filePath));

                    // Save the encrypted workbook to the target location using XLS format
                    workbook.Save(targetPath, SaveFormat.Excel97To2003);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Encryption completed for all XLS files.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
