// Title: Batch encrypt all .xls and .xlsx files in a folder with a shared password using Aspose.Cells for .NET (AES‑128)
// AI Prompts: Write C# code that scans a directory, loads each .xls or .xlsx workbook with Aspose.Cells, assigns a common password, and saves it as an AES‑128 protected .xlsx file. | Add functionality to the batch encryption script to write a log file containing the source path, destination path, and any errors for each processed workbook. | Refactor the program to accept the password and encryption strength (AES‑128 or AES‑256) from command‑line arguments.
// Common Searches: aspocells c# batch encrypt excel files in a folder with same password | how to apply AES-128 protection to multiple workbooks using Aspose.Cells .NET | convert xls to xlsx and set password programmatically with Aspose.Cells | c# encrypt all excel files in a directory using Aspose.Cells library | batch protect Excel workbooks with shared password Aspose.Cells example
// Tags: batch workbook encryption Aspose.Cells C# | shared password Excel files Aspose | AES-128 protection Aspose.Cells | convert xls to xlsx with password Aspose | directory iteration Excel encryption C#

using System;
using System.IO;
using Aspose.Cells;

// The example scans a specified folder for .xls and .xlsx files, loads each workbook with Aspose.Cells, applies a common password (AES‑128 encryption by default), and saves the protected workbook as a .xlsx file in an output directory, handling errors individually.
class BatchEncryptExcel
{
    static void Main()
    {
        // Directory containing the Excel files to encrypt
        string sourceDirectory = @"C:\ExcelFiles";

        // Directory where encrypted files will be saved (can be the same as source)
        string outputDirectory = @"C:\EncryptedExcelFiles";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputDirectory);

        // Shared password for all files
        string sharedPassword = "YourSecurePassword";

        // Get all files in the source directory
        string[] files = Directory.GetFiles(sourceDirectory, "*.*", SearchOption.TopDirectoryOnly);

        foreach (string filePath in files)
        {
            try
            {
                string extension = Path.GetExtension(filePath).ToLowerInvariant();
                if (extension != ".xls" && extension != ".xlsx")
                    continue; // Skip non‑Excel files

                // Verify the source file exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"Source file not found: {filePath}");
                    continue;
                }

                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Apply password protection (AES‑128 by default)
                workbook.Settings.Password = sharedPassword;

                // Build the output file path (preserve original name, force .xlsx)
                string outputFilePath = Path.Combine(
                    outputDirectory,
                    Path.GetFileNameWithoutExtension(filePath) + ".xlsx");

                // Save the workbook with the applied password
                workbook.Save(outputFilePath, SaveFormat.Xlsx);

                Console.WriteLine($"Encrypted: {Path.GetFileName(filePath)} -> {Path.GetFileName(outputFilePath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch encryption completed.");
    }
}
