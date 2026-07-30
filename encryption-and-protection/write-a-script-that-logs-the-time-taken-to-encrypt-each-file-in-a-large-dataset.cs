// Title: Log Encryption Time for Batch Excel Files Using Aspose.Cells (.NET C#)
// Description: A C# console utility that scans a folder for .xlsx workbooks, applies a password with strong 128‑bit encryption via Aspose.Cells, measures the duration of each encryption with Stopwatch, verifies the encrypted status, and writes the file name, encryption result, and elapsed milliseconds to the console.
// Keywords: Aspose.Cells encrypt Excel C# | measure encryption performance .NET | batch Excel encryption password | stopwatch encryption timing | strong cryptographic provider 128‑bit | detect encrypted workbook Aspose.Cells | C# Excel security automation | performance logging Excel files | file format detection Aspose.Cells | encrypt multiple xlsx files
// Common Searches: C# time Aspose.Cells workbook encryption | measure Excel file encryption duration .NET | batch encrypt xlsx with password Aspose.Cells | log encryption time per workbook C# | performance benchmark Aspose.Cells encryption | detect encrypted Excel file using Aspose.Cells
// Developer Intent: Encrypt each Excel workbook in a directory with a password and record how long the encryption takes for every file.
// Use Cases: Benchmark encryption speed before rolling out a security policy across thousands of spreadsheets. | Generate a compliance report that lists encryption status and processing time for each workbook. | Identify outlier files that cause slow encryption and target them for optimization.
// AI Prompts: Show how to export the timing results to a CSV file instead of the console. | Explain how to run the encryption loop in parallel while preserving per‑file timing data. | Demonstrate switching to 256‑bit encryption and logging the selected key size for each workbook.

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionTimer
{
    // A C# console utility that scans a folder for .xlsx workbooks, applies a password with strong 128‑bit encryption via Aspose.Cells, measures the duration of each encryption with Stopwatch, verifies the encrypted status, and writes the file name, encryption result, and elapsed milliseconds to the console.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the source Excel files
            string sourceFolder = @"C:\Data\ExcelFiles";
            // Folder where encrypted files will be saved
            string outputFolder = @"C:\Data\EncryptedFiles";

            // Verify source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder not found: {sourceFolder}");
                return;
            }

            // Ensure output folder exists
            Directory.CreateDirectory(outputFolder);

            // Password to protect the encrypted workbooks
            string password = "SecurePassword123";

            // Get all Excel files in the source folder (top‑level only)
            string[] files = Directory.GetFiles(sourceFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

            if (files.Length == 0)
            {
                Console.WriteLine("No Excel files found to encrypt.");
                return;
            }

            foreach (string filePath in files)
            {
                try
                {
                    // Verify the file still exists before processing
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found (skipped): {filePath}");
                        continue;
                    }

                    // Prepare output file path
                    string fileName = Path.GetFileNameWithoutExtension(filePath);
                    string encryptedFilePath = Path.Combine(outputFolder, $"{fileName}_encrypted.xlsx");

                    // Start timing the encryption process
                    Stopwatch sw = Stopwatch.StartNew();

                    // Load the workbook (no password needed for unencrypted source)
                    Workbook workbook = new Workbook(filePath);

                    // Set password protection
                    workbook.Settings.Password = password;

                    // Set encryption options (Strong encryption, 128‑bit key)
                    workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

                    // Save the encrypted workbook
                    workbook.Save(encryptedFilePath, SaveFormat.Xlsx);

                    // Stop timing
                    sw.Stop();

                    // Verify encryption status using FileFormatInfo
                    FileFormatInfo info = FileFormatUtil.DetectFileFormat(encryptedFilePath);
                    bool isEncrypted = info.IsEncrypted;

                    // Log the result
                    Console.WriteLine($"File: {Path.GetFileName(filePath)}");
                    Console.WriteLine($"Encrypted: {isEncrypted}");
                    Console.WriteLine($"Time taken: {sw.ElapsedMilliseconds} ms");
                    Console.WriteLine(new string('-', 40));
                }
                catch (Exception ex)
                {
                    // Log any error for the current file and continue with the next one
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Encryption processing completed.");
        }
    }
}
