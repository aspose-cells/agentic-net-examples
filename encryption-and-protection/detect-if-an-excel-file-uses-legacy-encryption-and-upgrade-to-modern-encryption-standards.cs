// Title: Upgrade Legacy Excel Encryption to AES‑128 with Aspose.Cells for .NET
// Description: Detects password‑protected legacy .xls files, loads them using the current password, applies a new password with strong AES‑128 encryption, and saves the workbook as a modern .xlsx file. Includes format detection, error handling, and automatic folder creation.
// Keywords: Aspose.Cells C# | detect encrypted Excel file | legacy .xls encryption | upgrade to AES 128 | convert .xls to .xlsx | set workbook password | strong cryptographic provider | FileFormatUtil DetectFileFormat | modern Excel protection | encryption upgrade .NET
// Common Searches: Aspose.Cells detect encrypted workbook | upgrade legacy .xls encryption to AES | change Excel file password using C# | convert password‑protected .xls to .xlsx | apply strong encryption to Excel with Aspose.Cells
// Developer Intent: Identify if an Excel workbook uses legacy encryption and re‑save it with AES‑128 protection and a new password.
// Use Cases: Migrate old password‑protected .xls reports to .xlsx with AES‑128 to satisfy security policies. | Batch‑process encrypted workbooks, replace weak passwords, and enforce modern encryption during file conversion. | Integrate encryption upgrade into automated data pipelines that handle legacy Excel assets.
// AI Prompts: Generate C# code using Aspose.Cells that checks whether an Excel file is legacy encrypted, opens it with the existing password, sets a new password, applies AES‑128 encryption, and saves as .xlsx. | Explain how to catch and handle exceptions when loading a password‑protected workbook with Aspose.Cells and verify the encryption type before saving. | Provide a step‑by‑step tutorial for upgrading legacy encrypted .xls files to modern encrypted .xlsx files in a .NET application.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionUpgrade
{
    // Detects password‑protected legacy .xls files, loads them using the current password, applies a new password with strong AES‑128 encryption, and saves the workbook as a modern .xlsx file. Includes format detection, error handling, and automatic folder creation.
    class Program
    {
        static void Main(string[] args)
        {
            // Example usage:
            // sourcePath - path to the original Excel file (could be .xls, .xlsx, etc.)
            // password    - current password required to open the file (null if not needed)
            // newPassword - password to protect the file with after upgrade
            // destPath    - path where the upgraded file will be saved (preferably .xlsx)

            string sourcePath = "legacy_encrypted.xls";
            string password = "oldPassword";
            string newPassword = "newStrongPassword";
            string destPath = "upgraded_encrypted.xlsx";

            try
            {
                UpgradeLegacyEncryption(sourcePath, password, newPassword, destPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        static void UpgradeLegacyEncryption(string sourcePath, string currentPassword, string newPassword, string destinationPath)
        {
            // Verify source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Detect file format and encryption status
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(sourcePath);

            // If the file is not encrypted, simply copy or inform the user
            if (!formatInfo.IsEncrypted)
            {
                Console.WriteLine("File is not encrypted. No upgrade needed.");
                // Optionally copy the file as‑is
                // File.Copy(sourcePath, destinationPath, true);
                return;
            }

            // Determine if the encryption is legacy (Excel 97‑2003 format) by file extension
            bool isLegacy = Path.GetExtension(sourcePath).Equals(".xls", StringComparison.OrdinalIgnoreCase);

            if (!isLegacy)
            {
                Console.WriteLine("File is encrypted but already uses a modern format. Re‑saving with new password.");
            }
            else
            {
                Console.WriteLine("Legacy encrypted file detected. Upgrading to modern encryption.");
            }

            // Load the workbook using the existing password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto)
            {
                Password = currentPassword // required for encrypted files
            };

            Workbook workbook;
            try
            {
                workbook = new Workbook(sourcePath, loadOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Apply new password and modern encryption options
            workbook.Settings.Password = newPassword; // set password for opening the file
            // Use strong encryption (AES 128‑bit) – EncryptionType is ignored for .xlsx but required for .xls
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Ensure destination directory exists
            string destDir = Path.GetDirectoryName(destinationPath);
            if (!string.IsNullOrEmpty(destDir) && !Directory.Exists(destDir))
            {
                Directory.CreateDirectory(destDir);
            }

            // Save in a modern format (XLSX) to ensure up‑to‑date encryption standards
            try
            {
                workbook.Save(destinationPath, SaveFormat.Xlsx);
                Console.WriteLine($"File has been upgraded and saved to '{destinationPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save upgraded workbook: {ex.Message}");
            }
        }
    }
}
