// Title: Convert a legacy‑encrypted .xls to AES‑256 protected .xlsx using Aspose.Cells for .NET
// Description: This example demonstrates how to open a password‑protected .xls workbook that uses an old XOR‑style encryption, assign a new password, switch to the StrongCryptographicProvider with a 256‑bit key, and save the file as an .xlsx with AES‑256 protection. The code uses LoadOptions for the original password and Workbook.SetEncryptionOptions for the upgrade, helping meet modern security and compliance requirements (e.g., GDPR, US SOX).
// Keywords: Aspose.Cells | C# | AES-256 encryption | StrongCryptographicProvider | legacy Excel encryption | XOR encrypted .xls | Workbook.SetEncryptionOptions | upgrade workbook encryption | password protected Excel | convert .xls to .xlsx | encryption migration | GDPR compliance
// Common Searches: how to change old XOR encrypted Excel file to AES‑256 with Aspose.Cells | Aspose.Cells upgrade legacy encryption to StrongCryptographicProvider | C# load encrypted .xls and save as .xlsx with AES‑256 | convert Excel97To2003 password protection to modern encryption | batch upgrade legacy encrypted workbooks Aspose.Cells
// Developer Intent: Migrate a workbook protected with legacy encryption to strong AES‑256 encryption.
// Use Cases: Modernize archived .xls files that use weak XOR encryption for regulatory compliance. | Automate bulk conversion of password‑protected legacy workbooks to AES‑256 secured .xlsx files. | Replace an outdated password scheme with a new strong password while preserving workbook data.
// AI Prompts: Generate C# code that opens a legacy‑encrypted .xls with a password, re‑encrypts it using AES‑256, and saves it as .xlsx using Aspose.Cells. | Explain the role of LoadOptions and SetEncryptionOptions when upgrading Excel workbook encryption in Aspose.Cells. | Create a reusable method: (string srcPath, string oldPwd, string newPwd, string destPath) → AES‑256 encrypted workbook.

using System;
using System.IO;
using Aspose.Cells;

// This example demonstrates how to open a password‑protected .xls workbook that uses an old XOR‑style encryption, assign a new password, switch to the StrongCryptographicProvider with a 256‑bit key, and save the file as an .xlsx with AES‑256 protection. The code uses LoadOptions for the original password and Workbook.SetEncryptionOptions for the upgrade, helping meet modern security and compliance requirements (e.g., GDPR, US SOX).
class UpgradeEncryption
{
    static void Main()
    {
        // Path to the workbook that uses legacy encryption (e.g., XOR, compatible, etc.)
        string legacyFilePath = "legacy_encrypted.xls";

        // Password required to open the legacy encrypted workbook
        string legacyPassword = "oldPassword";

        try
        {
            // Ensure the legacy file exists; if not, create a simple workbook with legacy encryption
            if (!File.Exists(legacyFilePath))
            {
                var tempWorkbook = new Workbook();
                tempWorkbook.Worksheets[0].Cells["A1"].PutValue("Sample data");
                tempWorkbook.Settings.Password = legacyPassword; // legacy encryption for .xls
                tempWorkbook.Save(legacyFilePath, SaveFormat.Excel97To2003);
            }

            // Load the legacy encrypted workbook using LoadOptions with the password
            var loadOptions = new LoadOptions
            {
                Password = legacyPassword
            };
            var workbook = new Workbook(legacyFilePath, loadOptions);

            // Define a new password for the upgraded workbook (can be the same or different)
            string newPassword = "newStrongPassword";

            // Assign the new password to the workbook settings
            workbook.Settings.Password = newPassword;

            // Upgrade encryption to AES‑256 (StrongCryptographicProvider with 256‑bit key)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

            // Save the workbook with the upgraded encryption
            string upgradedFilePath = "upgraded_encrypted.xlsx";
            workbook.Save(upgradedFilePath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook upgraded and saved to '{upgradedFilePath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
