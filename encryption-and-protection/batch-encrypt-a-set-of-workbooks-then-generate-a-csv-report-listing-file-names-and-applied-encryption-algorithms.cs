using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;

namespace BatchEncryptionExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input folder containing the workbooks to encrypt
            string inputFolder = @"C:\InputWorkbooks";
            // Output folder where encrypted workbooks will be saved
            string outputFolder = @"C:\EncryptedWorkbooks";
            // Path for the CSV report
            string reportPath = Path.Combine(outputFolder, "EncryptionReport.csv");
            // Password to protect all workbooks
            string password = "StrongPassword123";
            // Encryption algorithm to apply (will be recorded in the report)
            EncryptionType encryptionAlgorithm = EncryptionType.StrongCryptographicProvider;
            // Key length for encryption (128, 256, etc.)
            int keyLength = 128;

            // Ensure output directory exists
            Directory.CreateDirectory(outputFolder);

            // Prepare a list to hold CSV lines
            List<string> csvLines = new List<string>();
            // Add CSV header
            csvLines.Add("FileName,EncryptionAlgorithm");

            // Process each workbook file in the input folder (supports .xlsx and .xls)
            foreach (string filePath in Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly))
            {
                string extension = Path.GetExtension(filePath).ToLowerInvariant();
                if (extension != ".xlsx" && extension != ".xls" && extension != ".xlsm" && extension != ".xlsb")
                {
                    // Skip non‑Excel files
                    continue;
                }

                // Load the workbook (no password needed for original files)
                Workbook workbook = new Workbook(filePath);

                // Set the password that will be required to open the workbook
                workbook.Settings.Password = password;

                // Apply encryption options (algorithm and key length)
                workbook.SetEncryptionOptions(encryptionAlgorithm, keyLength);

                // Determine output file name (same name, placed in output folder)
                string outputFilePath = Path.Combine(outputFolder, Path.GetFileName(filePath));

                // Save the encrypted workbook
                workbook.Save(outputFilePath);

                // Record the file name and encryption algorithm used
                csvLines.Add($"{Path.GetFileName(filePath)},{encryptionAlgorithm}");

                // Dispose the workbook to free resources
                workbook.Dispose();
            }

            // Write the CSV report
            File.WriteAllLines(reportPath, csvLines);

            Console.WriteLine("Batch encryption completed.");
            Console.WriteLine($"Report generated at: {reportPath}");
        }
    }
}