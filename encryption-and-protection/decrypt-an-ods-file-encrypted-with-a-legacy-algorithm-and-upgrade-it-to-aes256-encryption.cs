// Title: Decrypt a legacy‑encrypted ODS and re‑encrypt it with AES‑256 using Aspose.Cells for .NET
// Description: Loads an ODS workbook protected with an old encryption algorithm via LoadOptions, removes the legacy protection, applies a new strong password and AES‑256 encryption with SetEncryptionOptions, and saves the file using OdsSaveOptions.
// Keywords: Aspose.Cells ODS decryption | legacy encryption upgrade | AES-256 ODS protection | LoadOptions password .NET | SetEncryptionOptions C# | OdsSaveOptions encryption | C# spreadsheet security | document encryption migration
// Common Searches: how to open a password protected ODS encrypted with legacy algorithm in Aspose.Cells | convert legacy ODS encryption to AES‑256 C# | Aspose.Cells change ODS password and encryption type | upgrade ODS file security .NET | re‑encrypt ODS with strong cryptography using Aspose
// Developer Intent: Read a password‑protected ODS encrypted with an outdated algorithm, then save it again secured with AES‑256 and a new password.
// Use Cases: Modernize archived ODS documents to comply with current data‑protection regulations. | Batch‑process a repository of ODS files, replacing weak passwords with strong AES‑256 encryption. | Integrate encryption renewal into a document management system to enforce consistent security policies.
// AI Prompts: Write C# code that opens a legacy‑encrypted ODS, removes its old protection, and saves it with AES‑256 using Aspose.Cells. | Explain the configuration of LoadOptions and OdsSaveOptions for migrating ODS encryption from legacy to strong cryptographic provider. | Suggest robust error‑handling patterns for re‑encrypting ODS files with Aspose.Cells in a .NET service.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Ods;

namespace AsposeCellsEncryptionUpgrade
{
    // Loads an ODS workbook protected with an old encryption algorithm via LoadOptions, removes the legacy protection, applies a new strong password and AES‑256 encryption with SetEncryptionOptions, and saves the file using OdsSaveOptions.
    class Program
    {
        static void Main()
        {
            // Paths and passwords
            string inputFile = "legacy_encrypted.ods";   // ODS encrypted with legacy algorithm
            string oldPassword = "oldPassword";          // Password used for the legacy encryption
            string newPassword = "newStrongPassword";    // Password for the new AES‑256 encryption
            string outputFile = "upgraded_encrypted.ods";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file not found: {inputFile}");
                return;
            }

            try
            {
                // Load the encrypted ODS file using LoadOptions with the legacy password
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Ods)
                {
                    Password = oldPassword // Provide password to decrypt the legacy file
                };

                Workbook workbook = new Workbook(inputFile, loadOptions);

                // Apply new encryption settings (AES‑256) and set the new password
                workbook.Settings.Password = newPassword; // Set password required before saving
                workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256); // AES‑256

                // Save the workbook as ODS with the new encryption
                OdsSaveOptions saveOptions = new OdsSaveOptions();
                // Optional: set ODF version if needed
                // saveOptions.OdfStrictVersion = OpenDocumentFormatVersionType.Odf12;

                workbook.Save(outputFile, saveOptions);

                Console.WriteLine("File has been decrypted and re‑encrypted with AES‑256 successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
