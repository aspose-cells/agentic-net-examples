// Title: C# – Decrypt Legacy‑Encrypted ODS and Re‑encrypt with AES‑256 using Aspose.Cells for .NET
// Description: Loads an ODS workbook protected with a legacy password, removes the old protection, applies a new password with StrongCryptographicProvider (AES‑256) and saves the file, showing how to upgrade ODS encryption in .NET.
// Keywords: Aspose.Cells | C# ODS decryption | legacy ODS encryption | AES-256 ODS | StrongCryptographicProvider | SetEncryptionOptions | LoadOptions password | upgrade ODS security | Encrypt ODS file .NET | Workbook encryption C#
// Common Searches: open password‑protected ODS with Aspose.Cells C# | convert legacy ODS encryption to AES‑256 | Aspose.Cells StrongCryptographicProvider example | change ODS file password and encryption in .NET | upgrade encrypted ODS files programmatically
// Developer Intent: Upgrade an ODS workbook from legacy password protection to AES‑256 encryption using C#.
// Use Cases: Modernize archived ODS spreadsheets to meet current security standards. | Batch‑process a repository of legacy‑encrypted ODS files, applying a new password and AES‑256 protection. | Integrate encryption upgrade into a document‑management pipeline that receives ODS files with outdated protection.
// AI Prompts: Write C# code that opens a password‑protected ODS file with Aspose.Cells and saves it using AES‑256 encryption. | Explain how to detect and handle an incorrect legacy password when re‑encrypting an ODS workbook. | Provide a step‑by‑step guide to batch convert a folder of ODS files from legacy encryption to StrongCryptographicProvider 256‑bit.

using System;
using System.IO;
using Aspose.Cells;

// Loads an ODS workbook protected with a legacy password, removes the old protection, applies a new password with StrongCryptographicProvider (AES‑256) and saves the file, showing how to upgrade ODS encryption in .NET.
class UpgradeEncryption
{
    static void Main()
    {
        // Paths to the source (legacy encrypted) and destination (AES‑256 encrypted) ODS files
        string sourcePath = "legacy_encrypted.ods";
        string destinationPath = "upgraded_encrypted.ods";

        // Passwords for the legacy file and the new encrypted file
        string legacyPassword = "oldPassword";
        string newPassword = "newPassword";

        try
        {
            // Verify source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the legacy ODS file using LoadOptions with the original password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Ods)
            {
                Password = legacyPassword
            };
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Apply a new password and set AES‑256 encryption (StrongCryptographicProvider, 256‑bit key)
            workbook.Settings.Password = newPassword;
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

            // Save the workbook as ODS with the upgraded encryption
            workbook.Save(destinationPath, SaveFormat.Ods);
            Console.WriteLine($"File saved successfully to {destinationPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
