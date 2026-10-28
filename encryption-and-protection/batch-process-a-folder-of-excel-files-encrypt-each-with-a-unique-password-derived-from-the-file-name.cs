// Title: Encrypt a batch of Excel workbooks in a folder using filename‑derived passwords with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells encrypt every .xls and .xlsx file in a directory using the file name as the password | How to apply password protection to multiple Excel workbooks while preserving their original format with Aspose.Cells | Batch set Workbook.Settings.Password based on file name in a Windows folder using Aspose.Cells | Automate encryption of Excel files in a folder with unique passwords derived from filenames in .NET | Save encrypted Excel workbook in its original format after setting password with Aspose.Cells
// Tags: Aspose.Cells workbook password encryption | filename derived Excel password .NET | preserve original SaveFormat when saving encrypted workbook | batch process .xls and .xlsx files C# | Workbook.Settings.Password programmatic setting | encrypt Excel files in folder Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example scans a specified directory for .xls, .xlsx, .xlsm, and .xlsb files, loads each workbook with Aspose.Cells, generates a password from the file name (without extension), assigns it to Workbook.Settings.Password, determines the original file format via a helper method, and saves the workbook back in the same format, overwriting the original file.
class ExcelBatchEncryptor
{
    static void Main()
    {
        // Path to the folder containing Excel files
        string folderPath = @"C:\Path\To\Folder";

        // Verify that the folder exists
        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Folder not found: {folderPath}");
            return;
        }

        string[] excelFiles;
        try
        {
            // Get all Excel files (both .xls and .xlsx) in the folder
            excelFiles = Directory.GetFiles(folderPath, "*.xls*");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error accessing files in folder: {ex.Message}");
            return;
        }

        foreach (string filePath in excelFiles)
        {
            // Ensure the file still exists before processing
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found (skipped): {filePath}");
                continue;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Derive a unique password from the file name (without extension)
                string password = Path.GetFileNameWithoutExtension(filePath);

                // Apply the password to the workbook (this encrypts the file)
                workbook.Settings.Password = password;

                // Determine the original file format to preserve it on save
                SaveFormat saveFormat = GetSaveFormat(Path.GetExtension(filePath));

                // Save the encrypted workbook, overwriting the original file
                workbook.Save(filePath, saveFormat);

                Console.WriteLine($"Encrypted '{Path.GetFileName(filePath)}' with password '{password}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{Path.GetFileName(filePath)}': {ex.Message}");
            }
        }
    }

    // Helper method to map file extensions to Aspose.Cells SaveFormat values
    private static SaveFormat GetSaveFormat(string extension)
    {
        switch (extension.ToLower())
        {
            case ".xls":
                return SaveFormat.Excel97To2003;
            case ".xlsx":
                return SaveFormat.Xlsx;
            case ".xlsm":
                return SaveFormat.Xlsm;
            case ".xlsb":
                return SaveFormat.Xlsb;
            default:
                // Default to XLSX if the extension is unrecognized
                return SaveFormat.Xlsx;
        }
    }
}
