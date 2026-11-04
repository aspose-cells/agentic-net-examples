// Title: Encrypt a batch of Excel files with individual passwords derived from file names using Aspose.Cells for .NET
// Common Searches: C# encrypt all .xlsx files in a folder with Aspose.Cells and unique passwords | How to apply a different password to each Excel workbook based on its filename using Aspose.Cells | Batch process Excel workbooks and save them encrypted while preserving original format in .NET | Generate password from file name and protect Excel files with Aspose.Cells C# example | Encrypt multiple Excel files programmatically with Aspose.Cells and custom password logic
// Tags: batch encrypt Excel workbooks Aspose.Cells | set workbook password Aspose.Cells C# | derive password from filename C# | save encrypted workbook preserving format Aspose.Cells | process directory of Excel files C#

using System;
using System.IO;
using Aspose.Cells;

// The program scans a specified input directory for .xls, .xlsx, and .xlsm files, generates a unique password from each file's name, applies the password to the workbook using Aspose.Cells, and saves the encrypted files to an output folder while keeping the original file format.
class Program
{
    static void Main(string[] args)
    {
        // Folder containing the source Excel files
        string sourceFolder = @"C:\InputExcel";

        // Folder where encrypted files will be saved
        string outputFolder = @"C:\EncryptedExcel";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Retrieve all Excel files (xls, xlsx, xlsm) in the source folder
        string[] files = Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly);
        foreach (string filePath in files)
        {
            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            if (ext != ".xls" && ext != ".xlsx" && ext != ".xlsm")
                continue; // Skip non‑Excel files

            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Derive a unique password from the file name (without extension)
            string fileName = Path.GetFileNameWithoutExtension(filePath);
            string password = GeneratePassword(fileName);

            // Apply the password to the workbook (encryption)
            workbook.Settings.Password = password;

            // Save the encrypted workbook, preserving the original format
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
            workbook.Save(outputPath, SaveFormat.Auto);
        }
    }

    // Simple password generator – can be replaced with any custom algorithm
    static string GeneratePassword(string baseName)
    {
        // Example: reverse the base name and append a constant suffix
        char[] chars = baseName.ToCharArray();
        Array.Reverse(chars);
        return new string(chars) + "Secure!";
    }
}
