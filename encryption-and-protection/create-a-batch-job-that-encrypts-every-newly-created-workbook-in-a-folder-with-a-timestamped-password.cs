// Title: C# batch program to encrypt newly created Excel .xlsx workbooks with a timestamp‑based password using Aspose.Cells
// Common Searches: how to add an open password to each .xlsx file in a folder using Aspose.Cells C# | C# script to encrypt multiple Excel workbooks with a generated timestamp password | batch encrypt Excel files with Aspose.Cells and save to another directory | Aspose.Cells encrypt workbook programmatically with dynamic password | automate Excel workbook protection in .NET using timestamped passwords
// Tags: Aspose.Cells encrypt workbook programmatically | C# timestamp password for Excel files | batch processing of .xlsx encryption | Settings.Password open password Aspose.Cells | SaveFormat.Xlsx encrypted workbook output

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookEncryptionBatch
{
    // The C# console application scans a source folder for .xlsx files, loads each workbook with Aspose.Cells, creates a password based on the current timestamp (yyyyMMddHHmmss), assigns it to the workbook's open password via Settings.Password, and saves the encrypted copy to a destination folder using SaveFormat.Xlsx. It logs progress, creates the destination directory if needed, and handles missing files or runtime errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the workbooks to be encrypted
            string sourceFolder = @"C:\Workbooks\Source";
            // Folder where encrypted workbooks will be saved
            string destinationFolder = @"C:\Workbooks\Encrypted";

            // Ensure source and destination folders exist
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder does not exist: {sourceFolder}");
                return;
            }
            Directory.CreateDirectory(destinationFolder);

            // Get all .xlsx files in the source folder
            string[] files = Directory.GetFiles(sourceFolder, "*.xlsx");

            foreach (string filePath in files)
            {
                try
                {
                    // Verify the file still exists before loading
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found, skipping: {filePath}");
                        continue;
                    }

                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Generate a timestamped password (e.g., 20231027143055)
                    string timestampPassword = DateTime.Now.ToString("yyyyMMddHHmmss");

                    // Apply the password to encrypt the workbook (open password)
                    workbook.Settings.Password = timestampPassword;
                    // Note: Write protection can be added via other APIs if needed

                    // Build the output file path
                    string fileName = Path.GetFileName(filePath);
                    string encryptedPath = Path.Combine(destinationFolder, fileName);

                    // Save the encrypted workbook
                    workbook.Save(encryptedPath, SaveFormat.Xlsx);

                    Console.WriteLine($"Encrypted workbook saved: {encryptedPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }
        }
    }
}
