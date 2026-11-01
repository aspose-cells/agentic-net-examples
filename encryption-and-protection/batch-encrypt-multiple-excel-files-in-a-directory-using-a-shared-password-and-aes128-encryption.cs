// Title: Batch encrypt multiple Excel .xlsx files with a shared password using Aspose.Cells (AES‑128) in C#
// Common Searches: C# code to encrypt all .xlsx files in a folder with the same password using Aspose.Cells | How to apply AES‑128 password protection to several Excel workbooks programmatically | Batch processing Excel files for encryption with Aspose.Cells library | Encrypt a directory of Excel files with a shared password in .NET | Save encrypted Excel workbooks to a different output folder using Aspose.Cells C#
// Tags: batch Excel encryption Aspose.Cells C# | AES‑128 password protection .xlsx Aspose.Cells | encrypt multiple workbooks programmatically | save encrypted workbook to output directory | process Excel files in folder Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample scans an input folder for .xlsx files, loads each workbook with Aspose.Cells, sets a shared password (AES‑128 encryption) via Workbook.Settings.Password, and saves the encrypted copies to a designated output folder.
class Program
{
    static void Main()
    {
        // Folder containing the original Excel files
        string inputFolder = @"C:\InputExcel";

        // Folder where encrypted files will be saved
        string outputFolder = @"C:\EncryptedExcel";

        // Shared password for all files
        string password = "SharedPassword";

        // Verify input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder does not exist: {inputFolder}");
            return;
        }

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Retrieve all .xlsx files in the input directory
        string[] excelFiles = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

        foreach (string filePath in excelFiles)
        {
            try
            {
                // Verify the source file exists
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"Source file not found: {filePath}");
                    continue;
                }

                // Load the workbook from the file
                Workbook workbook = new Workbook(filePath);

                // Set password protection (encryption)
                workbook.Settings.Password = password;

                // Determine the output file path
                string fileName = Path.GetFileName(filePath);
                string outputPath = Path.Combine(outputFolder, fileName);

                // Save the encrypted workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Encrypted file saved: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }
    }
}
