// Title: Batch encrypt Excel workbooks with Aspose.Cells for .NET and generate a CSV report
// Description: A C# console app that scans a folder, loads each .xls/.xlsx file with Aspose.Cells, applies a password and StrongCryptographicProvider (128‑bit) encryption, saves the protected copy to a target directory, and writes a CSV log of file names and the encryption algorithm used.
// Keywords: Aspose.Cells batch encryption .NET | C# encrypt multiple Excel files | Excel workbook password protection | StrongCryptographicProvider encryption | CSV audit log for encrypted workbooks | programmatic Excel security | bulk Excel file encryption | Aspose.Cells SetEncryptionOptions example
// Common Searches: how to encrypt all Excel files in a folder using Aspose.Cells | aspnet batch encrypt workbooks and create report | generate CSV list of encrypted Excel files | set encryption type and key length with Aspose.Cells | automate Excel password protection in C#
// Developer Intent: The developer needs to protect a collection of Excel workbooks with a uniform password and encryption algorithm, save the secured copies to a separate folder, and produce a CSV file that records each workbook’s name and the applied encryption method.
// Use Cases: Secure archival of legacy .xls and modern .xlsx reports before long‑term storage. | Create a compliance‑ready audit trail that lists encrypted files and their encryption settings. | Integrate workbook protection into a CI/CD pipeline that prepares Excel deliverables for external partners.
// AI Prompts: Refactor the batch encryption code to use async/await and "using" statements for better resource handling. | Show how to read a JSON configuration that maps individual files to different EncryptionType values and apply them during processing. | Demonstrate how to verify the encryption algorithm of a saved workbook using Aspose.Cells after encryption.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace BatchEncryptionExample
{
    // A C# console app that scans a folder, loads each .xls/.xlsx file with Aspose.Cells, applies a password and StrongCryptographicProvider (128‑bit) encryption, saves the protected copy to a target directory, and writes a CSV log of file names and the encryption algorithm used.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the source workbooks
            string sourceFolder = @"C:\Workbooks\Source";
            // Folder where encrypted workbooks will be saved
            string outputFolder = @"C:\Workbooks\Encrypted";

            // Ensure the output folder exists
            Directory.CreateDirectory(outputFolder);

            // CSV report file path
            string reportPath = Path.Combine(outputFolder, "EncryptionReport.csv");

            // Prepare a StringBuilder for the CSV content
            StringBuilder csvBuilder = new StringBuilder();
            csvBuilder.AppendLine("FileName,EncryptionAlgorithm");

            // Define the password and encryption settings to apply to all workbooks
            string password = "SecurePassword123";
            EncryptionType encryptionType = EncryptionType.StrongCryptographicProvider;
            int keyLength = 128; // 40, 128, or 256 bits are allowed

            // Process each workbook file in the source folder (supports .xlsx, .xls, .xlsm, etc.)
            foreach (string filePath in Directory.GetFiles(sourceFolder))
            {
                // Load the workbook (lifecycle: load)
                Workbook workbook = new Workbook(filePath);

                // Apply password protection
                workbook.Settings.Password = password;

                // Apply encryption options (the EncryptionType is ignored for 2007+ files but required by the API)
                workbook.SetEncryptionOptions(encryptionType, keyLength);

                // Determine the output file path (same name, different folder)
                string outputFilePath = Path.Combine(outputFolder, Path.GetFileName(filePath));

                // Save the encrypted workbook (lifecycle: save)
                workbook.Save(outputFilePath);

                // Record the result in the CSV report
                csvBuilder.AppendLine($"{Path.GetFileName(filePath)},{encryptionType}");

                // Dispose the workbook to free resources
                workbook.Dispose();
            }

            // Write the CSV report to disk
            File.WriteAllText(reportPath, csvBuilder.ToString());

            Console.WriteLine("Batch encryption completed. Report generated at:");
            Console.WriteLine(reportPath);
        }
    }
}
