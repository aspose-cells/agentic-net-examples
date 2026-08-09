// Title: Upgrade Legacy Encrypted Excel Files to AES‑128 with Aspose.Cells for .NET
// Description: Detect an Excel workbook’s legacy encryption using FileFormatUtil, open it with the original password, re‑apply StrongCryptographicProvider (AES‑128) via SetEncryptionOptions, and save the file as a modern .xlsx. The routine works for both encrypted and unencrypted sources and optionally updates the password.
// Keywords: Aspose.Cells | C# | .NET | legacy Excel encryption | AES‑128 re‑encryption | StrongCryptographicProvider | FileFormatUtil DetectFileFormat | LoadOptions password | SetEncryptionOptions | upgrade XLS to XLSX | password‑protected workbook
// Common Searches: detect legacy encryption in Excel with Aspose.Cells | convert old .xls password protection to AES .xlsx | re‑encrypt Excel file using StrongCryptographicProvider | upgrade encrypted workbook programmatically C# | Aspose.Cells change Excel file password and format
// Developer Intent: Identify legacy‑encrypted Excel files and re‑save them with modern AES encryption using Aspose.Cells.
// Use Cases: Batch migration of .xls files encrypted with the old algorithm to AES‑protected .xlsx files. | Opening a password‑protected legacy workbook, assigning a new password, and storing it securely. | Automated pipeline that validates incoming Excel files, upgrades encryption when needed, and archives the protected versions.
// AI Prompts: Generate C# code that uses Aspose.Cells to detect legacy encryption in an Excel file and upgrade it to AES‑128 with StrongCryptographicProvider. | Explain how LoadOptions and SetEncryptionOptions work together to re‑encrypt a workbook and change its password. | Create robust error‑handling for incorrect passwords, non‑encrypted files, and unsupported formats during encryption upgrade.

using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionUpgrade
{
    // Detect an Excel workbook’s legacy encryption using FileFormatUtil, open it with the original password, re‑apply StrongCryptographicProvider (AES‑128) via SetEncryptionOptions, and save the file as a modern .xlsx. The routine works for both encrypted and unencrypted sources and optionally updates the password.
    public class EncryptionUpgrade
    {
        /// <param name="inputPath">Path to the source Excel file.</param>
        /// <param name="password">Password required to open the encrypted file (null if not encrypted).</param>
        /// <param name="outputPath">Path where the upgraded file will be saved (should use a modern format like .xlsx).</param>
        public static void UpgradeEncryption(string inputPath, string password, string outputPath)
        {
            // Detect file format and encryption status
            FileFormatInfo fileInfo = FileFormatUtil.DetectFileFormat(inputPath);
            Console.WriteLine($"File format: {fileInfo.FileFormatType}");
            Console.WriteLine($"Is encrypted (legacy detection): {fileInfo.IsEncrypted}");

            // If the file is not encrypted, simply save it in the desired modern format
            if (!fileInfo.IsEncrypted)
            {
                // Load normally (no password needed)
                Workbook wb = new Workbook(inputPath);
                // Save as modern format (XLSX) – this automatically uses the current encryption standards
                wb.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine("File was not encrypted. Saved directly in modern format.");
                return;
            }

            // Load the encrypted workbook using the provided password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
            loadOptions.Password = password; // password may be null; Aspose will throw if required
            Workbook encryptedWb = new Workbook(inputPath, loadOptions);
            Console.WriteLine($"Workbook loaded. IsEncrypted after load: {encryptedWb.Settings.IsEncrypted}");

            // Apply modern encryption options (StrongCryptographicProvider with 128‑bit key)
            encryptedWb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Optionally set a new password (reuse the same one or change as needed)
            encryptedWb.Settings.Password = password;

            // Save the workbook in a modern format (XLSX) which uses AES encryption
            encryptedWb.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Encrypted workbook upgraded and saved to '{outputPath}'.");
        }

        // Example usage
        public static void Main()
        {
            string sourceFile = "legacy_encrypted.xls";   // legacy encrypted file (XLS)
            string pwd = "oldPassword";                  // password for the legacy file
            string upgradedFile = "upgraded.xlsx";       // target modern file

            UpgradeEncryption(sourceFile, pwd, upgradedFile);
        }
    }
}
