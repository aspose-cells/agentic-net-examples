// Title: How to decrypt a legacy‑encrypted ODS workbook and save it unencrypted with Aspose.Cells for .NET (AES‑256 not supported)
// AI Prompts: Load a legacy‑encrypted ODS file using Aspose.Cells LoadOptions with the password, then save the workbook using OdsSaveOptions. | Implement error handling to check for missing source files and invalid passwords when decrypting an ODS workbook in C#. | Describe why the current Aspose.Cells API cannot apply AES‑256 encryption to ODS files and demonstrate the fallback of saving the file without protection.
// Common Searches: c# aspnet open password protected ods file using aspose.cells | aspnet load ods file encrypted with legacy algorithm | remove password protection from ods workbook in c# | aes-256 encryption support for ods files in aspose.cells | save ods workbook unencrypted with aspose.cells c#
// Tags: load options for legacy ODS decryption | ods workbook save options without encryption | aes-256 encryption unsupported for ODS in Aspose.Cells | c# verify source file existence before loading workbook | handle password errors with Aspose.Cells LoadOptions

using System;
using System.IO;
using Aspose.Cells;

// The example checks that the source ODS file exists, loads it with the legacy password via LoadOptions, creates a Workbook, and saves it using OdsSaveOptions. Because Aspose.Cells does not currently support encryption for ODS files, the output is saved without any protection, and status messages are written to the console.
class OdsEncryptionUpgrade
{
    static void Main()
    {
        // Paths to the source (legacy encrypted) and destination ODS files
        string sourceFile = "legacy_encrypted.ods";
        string destinationFile = "aes256_encrypted.ods";

        // Password used for the legacy encrypted file
        string legacyPassword = "oldPassword";

        // New password for the upgraded AES‑256 encrypted file (not supported for ODS, kept for reference)
        string newPassword = "newPassword";

        try
        {
            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourceFile))
            {
                Console.WriteLine($"Source file not found: {sourceFile}");
                return;
            }

            // Load the legacy encrypted ODS file using the legacy password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Ods)
            {
                // Legacy ODS encryption is password‑based; set the password here
                Password = legacyPassword
            };

            // Create a Workbook instance from the encrypted ODS
            Workbook workbook = new Workbook(sourceFile, loadOptions);

            // NOTE: Aspose.Cells does not currently support setting encryption for ODS files.
            // The workbook is saved without encryption. If future versions add support,
            // the appropriate properties can be set on OdsSaveOptions.

            // Prepare save options for ODS (no encryption settings available)
            OdsSaveOptions saveOptions = new OdsSaveOptions();

            // Save the workbook (unencrypted)
            workbook.Save(destinationFile, saveOptions);

            Console.WriteLine("File has been decrypted and saved (encryption not supported for ODS).");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
