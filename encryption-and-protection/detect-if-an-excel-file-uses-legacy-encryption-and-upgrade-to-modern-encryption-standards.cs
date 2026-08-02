// Title: Upgrade Legacy Excel Encryption to 256‑bit StrongCryptographicProvider with Aspose.Cells for .NET (C#)
// Description: Shows how to detect an Excel workbook encrypted with legacy protection, open it using the original password, re‑apply encryption with Aspose.Cells' StrongCryptographicProvider (256‑bit key), and save the file, while copying unencrypted workbooks unchanged.
// Keywords: Aspose.Cells | C# | legacy Excel encryption | StrongCryptographicProvider | 256-bit encryption | upgrade Excel password protection | detect encrypted workbook | re‑encrypt Excel file | Excel encryption .NET | convert .xls to .xlsx encryption
// Common Searches: How to detect legacy encryption in an Excel file using Aspose.Cells | Upgrade Excel .xls password protection to 256‑bit encryption C# | Re‑encrypt encrypted Excel workbook with StrongCryptographicProvider Aspose | Convert old encrypted Excel files to modern encryption .NET | Programmatically change Excel file encryption type Aspose.Cells
// Developer Intent: Programmatically identify Excel files protected with legacy encryption and re‑save them using Aspose.Cells with modern 256‑bit StrongCryptographicProvider encryption.
// Use Cases: Batch migration of archived .xls files to .xlsx with strong encryption | Compliance‑driven re‑encryption of confidential spreadsheets before storage | Automated data‑pipeline step that upgrades encryption of incoming Excel uploads | Secure backup creation for legacy workbooks | Standardizing encryption across corporate Excel assets
// AI Prompts: Generate C# code that checks if an Excel workbook is encrypted, opens it with a given password, and saves it with StrongCryptographicProvider 256‑bit encryption using Aspose.Cells. | Create a robust error‑handling wrapper for upgrading Excel file encryption, including handling of missing files and incorrect passwords. | Write a script that scans a directory, detects encrypted Excel files, and upgrades each to modern encryption while preserving the original password. | Provide unit tests for the encryption‑upgrade method using Aspose.Cells mock objects.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionUpgrade
{
    // Shows how to detect an Excel workbook encrypted with legacy protection, open it using the original password, re‑apply encryption with Aspose.Cells' StrongCryptographicProvider (256‑bit key), and save the file, while copying unencrypted workbooks unchanged.
    public class EncryptionUpgrade
    {
        public static void UpgradeLegacyEncryption(string inputPath, string password, string outputPath)
        {
            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Detect file format and encryption status
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(inputPath);
                Console.WriteLine($"IsEncrypted: {formatInfo.IsEncrypted}");

                // If the file is not encrypted, copy it to the output location
                if (!formatInfo.IsEncrypted)
                {
                    Console.WriteLine("File is not encrypted. No upgrade needed.");
                    // Ensure the output directory exists
                    Directory.CreateDirectory(Path.GetDirectoryName(outputPath) ?? string.Empty);
                    File.Copy(inputPath, outputPath, overwrite: true);
                    Console.WriteLine($"File copied to '{outputPath}'.");
                    return;
                }

                // Load the encrypted workbook using the supplied password
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto) { Password = password };
                Workbook workbook = new Workbook(inputPath, loadOptions);
                Console.WriteLine("Workbook loaded successfully.");

                // Apply modern encryption options (StrongCryptographicProvider with 256‑bit key)
                workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);
                // Set the password for the new encryption
                workbook.Settings.Password = password;

                // Ensure the output directory exists
                Directory.CreateDirectory(Path.GetDirectoryName(outputPath) ?? string.Empty);
                // Save the workbook with the upgraded encryption
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved with upgraded encryption to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during encryption upgrade: {ex.Message}");
            }
        }

        // Example usage
        public static void Run()
        {
            string sourceFile = "legacy_encrypted.xls";          // Path to the legacy encrypted file
            string pwd = "oldPassword";                         // Password used to open the legacy file
            string upgradedFile = "upgraded_encrypted.xlsx";    // Desired output file

            UpgradeLegacyEncryption(sourceFile, pwd, upgradedFile);
        }

        // Entry point required for console application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
