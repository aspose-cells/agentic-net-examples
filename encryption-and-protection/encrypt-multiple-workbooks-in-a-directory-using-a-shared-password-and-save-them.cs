// Title: Batch encrypt Excel workbooks with a shared password using Aspose.Cells for .NET (C#)
// Description: Iterates through all supported Excel files in a source folder, applies the same password via Workbook.Settings.Password, and saves encrypted copies to a target directory while preserving original filenames.
// Keywords: Aspose.Cells | C# | batch Excel encryption | Excel password protection | encrypt multiple workbooks | programmatic workbook security | save encrypted workbook | directory processing
// Common Searches: C# encrypt all Excel files in a folder with Aspose.Cells | batch apply password to workbooks using Aspose.Cells .NET | how to protect multiple Excel workbooks programmatically | save encrypted copies of Excel files to another directory
// Developer Intent: Apply a single password to every workbook in a directory and store the encrypted versions in a separate folder.
// Use Cases: Secure a batch of financial reports before archiving to meet data‑protection regulations. | Create password‑protected spreadsheets for safe upload to cloud storage or shared drives. | Distribute locked template workbooks to external partners without exposing source data.
// AI Prompts: Generate C# code that encrypts all Excel files in a folder using Aspose.Cells and a shared password, preserving original filenames. | Show how to modify the sample to generate a unique password for each workbook based on its filename or metadata. | Add comprehensive logging and exception handling to the batch encryption process.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    // Iterates through all supported Excel files in a source folder, applies the same password via Workbook.Settings.Password, and saves encrypted copies to a target directory while preserving original filenames.
    class Program
    {
        static void Main()
        {
            // Directory containing the workbooks to encrypt
            string inputDirectory = @"C:\InputWorkbooks";

            // Directory where encrypted workbooks will be saved
            string outputDirectory = @"C:\EncryptedWorkbooks";

            // Shared password for all workbooks
            string sharedPassword = "MySharedPassword";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputDirectory);

            // Process each Excel file in the input directory (supports common formats)
            foreach (string filePath in Directory.GetFiles(inputDirectory, "*.*"))
            {
                // Filter supported Excel extensions
                string extension = Path.GetExtension(filePath).ToLowerInvariant();
                if (extension != ".xlsx" && extension != ".xls" && extension != ".xlsm" && extension != ".xlsb" && extension != ".ods")
                    continue;

                // Load the workbook (lifecycle rule: load)
                Workbook workbook = new Workbook(filePath);

                // Set the password to encrypt the workbook (lifecycle rule: create/use Settings)
                workbook.Settings.Password = sharedPassword;

                // Build the output file path (overwrite the original name in the output folder)
                string outputPath = Path.Combine(outputDirectory, Path.GetFileName(filePath));

                // Save the encrypted workbook (lifecycle rule: save)
                workbook.Save(outputPath);

                // Release resources
                workbook.Dispose();
            }

            Console.WriteLine("Encryption of workbooks completed.");
        }
    }
}
