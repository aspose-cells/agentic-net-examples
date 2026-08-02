// Title: Upgrade Excel workbook encryption to AES‑256 with Aspose.Cells for .NET
// Description: Loads a password‑protected .xlsx using LoadOptions, retains the original password, switches the encryption to StrongCryptographicProvider (AES‑256), and saves the workbook as a new file with stronger protection. Includes file‑existence check and error handling.
// Keywords: Aspose.Cells | C# | Excel encryption | AES-256 | upgrade encryption | password protected workbook | LoadOptions | SetEncryptionOptions | .NET | StrongCryptographicProvider
// Common Searches: Aspose.Cells upgrade workbook encryption to AES 256 | change Excel file encryption algorithm .NET | load encrypted Excel with password Aspose.Cells | set encryption options StrongCryptographicProvider C# | re‑save password protected workbook with stronger encryption
// Developer Intent: Replace the legacy encryption of an existing password‑protected Excel file with AES‑256 and save the updated workbook.
// Use Cases: Modernize legacy Excel files to comply with current security standards. | Automate re‑encryption of uploaded workbooks to a uniform AES‑256 policy. | Refresh workbook protection after a password rotation without altering the password.
// AI Prompts: Write C# code that opens an encrypted .xlsx using Aspose.Cells, upgrades its encryption to AES‑256, and saves it as a new file. | Explain how to catch and log exceptions when changing workbook encryption while preserving the original password. | Provide steps to verify that a saved workbook is encrypted with AES‑256 using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Loads a password‑protected .xlsx using LoadOptions, retains the original password, switches the encryption to StrongCryptographicProvider (AES‑256), and saves the workbook as a new file with stronger protection. Includes file‑existence check and error handling.
class UpgradeWorkbookEncryption
{
    static void Main()
    {
        // Path to the existing encrypted workbook
        string inputPath = "EncryptedOld.xlsx";

        // Path where the upgraded encrypted workbook will be saved
        string outputPath = "EncryptedUpgraded.xlsx";

        // Existing password used to open the workbook
        string password = "oldPassword";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the encrypted workbook using the provided password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = password
            };
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Preserve the original password for saving
            workbook.Settings.Password = password;

            // Upgrade encryption to a stronger algorithm (AES 256‑bit)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

            // Save the workbook with the upgraded encryption
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook encrypted with upgraded settings saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
