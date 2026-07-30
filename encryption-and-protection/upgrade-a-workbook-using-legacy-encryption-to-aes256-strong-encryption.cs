// Title: Convert Legacy‑Encrypted Excel Workbook to AES‑256 Using Aspose.Cells for .NET
// Description: Load an Excel file protected with legacy encryption (e.g., XOR), apply a new password, upgrade the encryption to AES‑256 via StrongCryptographicProvider, save as .xlsx, and verify the encryption status—all with Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# | Excel encryption | AES-256 | StrongCryptographicProvider | legacy encryption upgrade | set workbook password | LoadOptions password | Workbook.SetEncryptionOptions | encrypt Excel file .NET
// Common Searches: How to upgrade an Excel file encrypted with XOR to AES‑256 using Aspose.Cells | Aspose.Cells C# change workbook encryption to StrongCryptographicProvider | Convert legacy encrypted .xls to .xlsx with AES‑256 | Set new password and encryption type for existing workbook Aspose.Cells | Verify Excel file encryption after saving with Aspose.Cells
// Developer Intent: Upgrade a workbook protected with legacy encryption to AES‑256 strong encryption via Aspose.Cells.
// Use Cases: Modernize archived reports encrypted with outdated algorithms for compliance | Re‑encrypt user‑uploaded spreadsheets before cloud storage | Automate batch conversion of legacy‑encrypted Excel files to AES‑256 | Validate encryption after password change in automated pipelines
// AI Prompts: Provide C# code that opens a legacy‑encrypted .xls with a password, switches to AES‑256 using StrongCryptographicProvider, and saves as .xlsx. | Show how to set a new password and specify 256‑bit encryption for an Aspose.Cells workbook. | Explain how to confirm that a saved workbook is encrypted and handle incorrect legacy passwords in Aspose.Cells. | Give a step‑by‑step guide to batch upgrade multiple Excel files from XOR to AES‑256 with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionUpgrade
{
    // Load an Excel file protected with legacy encryption (e.g., XOR), apply a new password, upgrade the encryption to AES‑256 via StrongCryptographicProvider, save as .xlsx, and verify the encryption status—all with Aspose.Cells for C#.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the legacy‑encrypted workbook (e.g., encrypted with XOR or compatible encryption)
                string legacyFilePath = "LegacyEncryptedWorkbook.xls";

                // Verify that the legacy file exists before attempting to load it
                if (!File.Exists(legacyFilePath))
                {
                    Console.WriteLine($"Error: Legacy file not found at path '{legacyFilePath}'.");
                    return;
                }

                // Password used to open the legacy encrypted workbook
                string legacyPassword = "oldPassword";

                // Load the legacy encrypted workbook using LoadOptions with the password
                LoadOptions loadOptions = new LoadOptions
                {
                    Password = legacyPassword
                };
                Workbook workbook = new Workbook(legacyFilePath, loadOptions);

                // Set a new password for the workbook (can be the same or a new one)
                string newPassword = "newStrongPassword";
                workbook.Settings.Password = newPassword;

                // Upgrade encryption to AES‑256 (StrongCryptographicProvider with 256‑bit key)
                workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

                // Save the workbook with the upgraded encryption
                string upgradedFilePath = "UpgradedEncryptedWorkbook.xlsx";
                workbook.Save(upgradedFilePath, SaveFormat.Xlsx);

                // Optional: verify that the workbook is now encrypted
                Workbook verifyWorkbook = new Workbook(upgradedFilePath, new LoadOptions { Password = newPassword });
                Console.WriteLine("Upgrade successful, IsEncrypted: " + verifyWorkbook.Settings.IsEncrypted);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
