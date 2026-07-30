// Title: Upgrade an Encrypted Excel Workbook to AES‑256 with Aspose.Cells for .NET
// Description: Shows how to open a password‑protected .xlsx file using Aspose.Cells LoadOptions, assign the same password, upgrade the encryption to AES‑256 via SetEncryptionOptions (StrongCryptographicProvider), and save the workbook with a new name.
// Keywords: Aspose.Cells | C# | Excel encryption | AES 256 | SetEncryptionOptions | StrongCryptographicProvider | upgrade workbook encryption | load encrypted workbook | password protected Excel | re‑encrypt Excel file | Aspose.Cells .NET
// Common Searches: Aspose.Cells change workbook encryption to AES 256 | C# load encrypted Excel and re‑save with stronger encryption | upgrade old Excel password protection using Aspose.Cells | SetEncryptionOptions example C# | how to re‑encrypt an Excel file with Aspose.Cells
// Developer Intent: Open a password‑protected workbook, replace its existing encryption with AES‑256, and write the updated file.
// Use Cases: Modernize legacy reports that use weak encryption before archiving for compliance. | Batch‑process a folder of workbooks to enforce AES‑256 protection across the organization. | Apply a new corporate password policy by re‑encrypting existing files with a stronger algorithm.
// AI Prompts: Generate C# code that opens an encrypted .xlsx using Aspose.Cells, upgrades the encryption to AES‑256 with StrongCryptographicProvider, and saves it as a new file. | Explain the parameters of Workbook.SetEncryptionOptions and list the EncryptionType values that correspond to AES‑256 in Aspose.Cells. | Create robust error‑handling for missing input files, incorrect passwords, and unsupported encryption types while upgrading workbook security.

using System;
using System.IO;
using Aspose.Cells;

namespace UpgradeWorkbookEncryptionApp
{
    // Shows how to open a password‑protected .xlsx file using Aspose.Cells LoadOptions, assign the same password, upgrade the encryption to AES‑256 via SetEncryptionOptions (StrongCryptographicProvider), and save the workbook with a new name.
    class UpgradeWorkbookEncryption
    {
        static void Main()
        {
            try
            {
                // Path to the existing encrypted workbook
                string inputFile = "EncryptedOld.xlsx";

                // Verify the input file exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Password used to open the existing encrypted workbook
                string password = "oldPassword";

                // Load the encrypted workbook with the password
                LoadOptions loadOptions = new LoadOptions
                {
                    Password = password
                };
                Workbook workbook = new Workbook(inputFile, loadOptions);

                // Set the same password for saving (required for encryption)
                workbook.Settings.Password = password;

                // Upgrade encryption to a stronger algorithm (AES 256‑bit)
                workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

                // Save the workbook with the upgraded encryption
                string outputFile = "EncryptedUpgraded.xlsx";
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
