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
            // Folder containing the source workbooks
            string inputFolder = @"C:\InputWorkbooks";

            // Folder where encrypted workbooks will be saved
            string outputFolder = @"C:\EncryptedWorkbooks";

            // Ensure the output folder exists
            Directory.CreateDirectory(outputFolder);

            // Password to protect the workbooks
            string password = "SecurePassword123";

            // Encryption settings (algorithm and key length)
            EncryptionType encryptionType = EncryptionType.StrongCryptographicProvider;
            int keyLength = 128; // Valid values: 40, 128, 256

            // List to hold CSV report lines
            List<string> reportLines = new List<string>();
            // Add CSV header
            reportLines.Add("FileName,EncryptionAlgorithm");

            // Process each Excel file in the input folder
            foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
            {
                // Load the workbook (lifecycle rule: load)
                using (Workbook workbook = new Workbook(filePath))
                {
                    // Apply password protection
                    workbook.Settings.Password = password;

                    // Apply encryption options (lifecycle rule: create -> set options)
                    workbook.SetEncryptionOptions(encryptionType, keyLength);

                    // Determine output file path
                    string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));

                    // Save the encrypted workbook (lifecycle rule: save)
                    workbook.Save(outputPath);

                    // Record the file name and encryption algorithm used
                    reportLines.Add($"{Path.GetFileName(filePath)},{encryptionType}");
                }
            }

            // Write the CSV report
            string reportPath = Path.Combine(outputFolder, "EncryptionReport.csv");
            File.WriteAllLines(reportPath, reportLines);

            Console.WriteLine("Batch encryption completed. Report generated at:");
            Console.WriteLine(reportPath);
        }
    }
}