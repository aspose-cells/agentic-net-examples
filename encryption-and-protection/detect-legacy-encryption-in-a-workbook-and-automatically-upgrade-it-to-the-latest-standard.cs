// Title: Detect legacy encrypted Excel workbook and upgrade to AES‑256 encryption with Aspose.Cells for .NET
// Description: This C# example shows how to use Aspose.Cells to detect whether an Excel file is password‑protected, open it with the supplied password, apply the StrongCryptographicProvider AES‑256 encryption, and save the workbook in the modern XLSX format. It also handles non‑encrypted files by simply converting them to XLSX.
// Keywords: Aspose.Cells | C# | Excel encryption upgrade | AES-256 | StrongCryptographicProvider | legacy .xls password | FileFormatUtil DetectFileFormat | LoadOptions password | SetEncryptionOptions | workbook protection | Excel security .NET | convert encrypted XLS to XLSX
// Common Searches: how to detect encrypted Excel file using Aspose.Cells | upgrade legacy .xls password protection to AES-256 in C# | Aspose.Cells set StrongCryptographicProvider encryption | convert encrypted workbook to XLSX with Aspose.Cells | C# code to re‑encrypt Excel file with stronger algorithm
// Developer Intent: Automatically detect a legacy encrypted Excel workbook, re‑encrypt it with AES‑256, and save as XLSX.
// Use Cases: Batch‑process incoming legacy .xls files, upgrade their protection to meet modern compliance standards. | Integrate encryption strengthening into a document‑management workflow without manual intervention. | Migrate password‑protected workbooks to a stronger algorithm before archiving or sharing.
// AI Prompts: Generate a C# method that uses Aspose.Cells to detect if a workbook is encrypted and upgrades it to AES‑256 with the same password. | Provide robust error handling for missing files, incorrect passwords, and unsupported formats when upgrading Excel encryption. | Create unit tests for the UpgradeWorkbookEncryption function covering encrypted, unencrypted, and error scenarios.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionUpgrade
{
    // This C# example shows how to use Aspose.Cells to detect whether an Excel file is password‑protected, open it with the supplied password, apply the StrongCryptographicProvider AES‑256 encryption, and save the workbook in the modern XLSX format. It also handles non‑encrypted files by simply converting them to XLSX.
    class Program
    {
        static void Main(string[] args)
        {
            // Example usage:
            // Input workbook path (could be legacy encrypted)
            string inputPath = "LegacyEncryptedWorkbook.xls";
            // Password required to open the legacy encrypted workbook
            string password = "oldPassword";
            // Path for the upgraded workbook
            string outputPath = "UpgradedWorkbook.xlsx";

            try
            {
                UpgradeWorkbookEncryption(inputPath, password, outputPath);
                Console.WriteLine("Workbook encryption upgrade completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        /// <param name="inputFilePath">Path to the source workbook.</param>
        /// <param name="password">Password to open the source workbook (null if not needed).</param>
        /// <param name="outputFilePath">Path where the upgraded workbook will be saved.</param>
        static void UpgradeWorkbookEncryption(string inputFilePath, string password, string outputFilePath)
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputFilePath))
            {
                throw new FileNotFoundException($"Input file not found: {inputFilePath}");
            }

            // Detect file format and encryption status without loading the whole workbook
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(inputFilePath);

            // If the file is not encrypted, simply load and save (no upgrade needed)
            if (!formatInfo.IsEncrypted)
            {
                Workbook wb = new Workbook(inputFilePath);
                wb.Save(outputFilePath, SaveFormat.Xlsx);
                return;
            }

            // The file is encrypted; load it using the provided password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = password
            };
            Workbook encryptedWb = new Workbook(inputFilePath, loadOptions);

            // Apply the latest encryption options (AES‑256)
            encryptedWb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

            // Keep the same password (or set a new one if desired)
            encryptedWb.Settings.Password = password;

            // Save the workbook in a modern format (XLSX) which embeds the upgraded encryption
            encryptedWb.Save(outputFilePath, SaveFormat.Xlsx);
        }
    }
}
