// Title: C# – Upgrade legacy‑encrypted Excel (.xls) to AES‑256 with Aspose.Cells
// Description: Loads a password‑protected .xls using Aspose.Cells, assigns a new password, switches encryption to StrongCryptographicProvider (256‑bit AES‑256), and saves the workbook as .xlsx with strong protection.
// Keywords: Aspose.Cells | C# | upgrade Excel encryption | AES-256 | StrongCryptographicProvider | legacy .xls password | SetEncryptionOptions | Workbook encryption conversion | secure Excel files | SaveFormat.Xlsx
// Common Searches: convert old .xls password protection to AES‑256 C# | Aspose.Cells change workbook encryption type | upgrade legacy Excel encryption with StrongCryptographicProvider | C# re‑encrypt Excel file using Aspose.Cells | how to set 256‑bit encryption for Excel workbook
// Developer Intent: Open a legacy‑encrypted workbook, apply a new password, and resave it with AES‑256 encryption.
// Use Cases: Modernize archived .xls files for compliance by applying 256‑bit encryption. | Replace weak workbook passwords before sharing files with external partners. | Automate batch conversion of legacy protected spreadsheets to a secure .xlsx format for cloud storage.
// AI Prompts: Write C# code that opens a password‑protected .xls with Aspose.Cells, changes the password, and saves it as .xlsx encrypted with AES‑256. | Show how to verify that a saved workbook uses StrongCryptographicProvider encryption and retrieve its encryption settings. | Provide robust error‑handling patterns for upgrading Excel file encryption using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace UpgradeEncryptionApp
{
    // Loads a password‑protected .xls using Aspose.Cells, assigns a new password, switches encryption to StrongCryptographicProvider (256‑bit AES‑256), and saves the workbook as .xlsx with strong protection.
    class UpgradeEncryption
    {
        static void Main()
        {
            try
            {
                // Path to the workbook that uses legacy encryption
                string inputPath = "legacy_encrypted.xls";

                // Password that unlocks the legacy encrypted workbook
                string legacyPassword = "oldPassword";

                // Password to be used for the upgraded encryption (can be the same or different)
                string newPassword = "newStrongPassword";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook with the legacy password
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
                loadOptions.Password = legacyPassword;
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Assign the password that will protect the workbook after re‑saving
                workbook.Settings.Password = newPassword;

                // Upgrade encryption to AES‑256 (StrongCryptographicProvider, 256‑bit key)
                workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

                // Save the workbook with the stronger encryption
                string outputPath = "upgraded_encrypted.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved with upgraded encryption: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
