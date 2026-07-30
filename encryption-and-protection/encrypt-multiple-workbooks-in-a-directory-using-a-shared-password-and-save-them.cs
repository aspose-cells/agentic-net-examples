// Title: Batch encrypt Excel workbooks in a folder with a shared password using Aspose.Cells for .NET
// Description: A C# console app that scans a given directory, loads each supported Excel file (.xlsx, .xls, .xlsm, .xlsb, .ods) with Aspose.Cells, sets a common password via Workbook.Settings.Password, overwrites the original file with the encrypted version, disposes resources, and writes a log entry for every processed workbook.
// Keywords: Aspose.Cells batch encryption | C# encrypt multiple Excel files | set shared password for workbooks | programmatic Excel file protection .NET | encrypt .xlsx .xls .xlsm .xlsb .ods | automated workbook security | Aspose.Cells password protection example
// Common Searches: how to encrypt all Excel files in a folder using Aspose.Cells | batch apply password to multiple workbooks C# | Aspose.Cells encrypt many spreadsheets with one password | protect .ods files programmatically .NET | automate Excel file encryption in a build pipeline
// Developer Intent: Apply the same password to every Excel workbook in a specified directory in a single automated run.
// Use Cases: Secure a batch of financial reports before distribution to auditors. | Integrate workbook encryption into a nightly CI/CD job to meet data‑privacy regulations. | Automatically protect user‑uploaded Excel files on a web server with a corporate password.
// AI Prompts: Generate C# code that uses Aspose.Cells to encrypt all Excel files in a folder with a shared password and prints the name of each encrypted file. | Add robust error handling to the batch encryption loop, logging skipped files and any exceptions that occur. | Modify the example to write encrypted copies to a separate output directory while keeping the original filenames unchanged.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    // A C# console app that scans a given directory, loads each supported Excel file (.xlsx, .xls, .xlsm, .xlsb, .ods) with Aspose.Cells, sets a common password via Workbook.Settings.Password, overwrites the original file with the encrypted version, disposes resources, and writes a log entry for every processed workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Directory containing the workbooks to encrypt
            string sourceDirectory = @"C:\InputWorkbooks";

            // Shared password for all workbooks
            string sharedPassword = "MySecretPassword";

            // Get all Excel files in the directory (including .xlsx, .xls, .xlsm, etc.)
            string[] workbookFiles = Directory.GetFiles(sourceDirectory, "*.*", SearchOption.TopDirectoryOnly);
            
            foreach (string filePath in workbookFiles)
            {
                // Filter only supported Excel extensions
                string extension = Path.GetExtension(filePath).ToLowerInvariant();
                if (extension != ".xlsx" && extension != ".xls" && extension != ".xlsm" && extension != ".xlsb" && extension != ".ods")
                {
                    continue; // Skip non‑Excel files
                }

                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Set the password to encrypt the workbook
                workbook.Settings.Password = sharedPassword;

                // Save the workbook (overwrites the original file with the encrypted version)
                workbook.Save(filePath);

                // Release resources
                workbook.Dispose();

                Console.WriteLine($"Encrypted workbook: {Path.GetFileName(filePath)}");
            }

            Console.WriteLine("All workbooks have been encrypted.");
        }
    }
}
