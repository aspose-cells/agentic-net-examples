// Title: Decrypt a legacy‑encrypted ODS workbook and attempt AES‑256 re‑encryption using Aspose.Cells for .NET
// Common Searches: c# aspose.cells load password protected ods workbook with legacy encryption | aspnet upgrade ods file encryption to aes‑256 using aspose.cells | how to use LoadOptions to open encrypted ODS in Aspose.Cells C# | does Aspose.Cells support AES‑256 encryption for ODS format | save ods workbook unencrypted after decrypting legacy password in C#
// Tags: loadoptions password ods aspnet | ods decryption with aspose.cells | aes-256 encryption upgrade ods aspnet | aspose.cells ods password handling | ods saveoptions unencrypted output | legacy ods encryption removal c#

using System;
using System.IO;
using Aspose.Cells;

// The example loads an ODS workbook encrypted with a legacy password using LoadOptions, then saves it with OdsSaveOptions. Current Aspose.Cells versions do not provide AES‑256 encryption for ODS, so the result is an unencrypted ODS file; future versions may add encryption support.
class OdsEncryptionUpgrade
{
    static void Main()
    {
        // Path to the original ODS file encrypted with a legacy algorithm
        string inputPath = @"C:\Files\LegacyEncrypted.ods";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        // Password used for the legacy encryption
        string legacyPassword = "legacyPwd";

        // New password for AES‑256 encryption (if supported)
        string newPassword = "newStrongPwd";

        // Path for the upgraded ODS file
        string outputPath = @"C:\Files\UpgradedEncrypted.ods";

        try
        {
            // Load the ODS file using the legacy password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Ods)
            {
                Password = legacyPassword
            };
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Prepare ODS save options.
            // Note: As of current Aspose.Cells versions, ODS format does not support
            // setting a password or encryption algorithm via OdsSaveOptions.
            OdsSaveOptions saveOptions = new OdsSaveOptions();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (unencrypted ODS). If future versions add encryption support,
            // the appropriate properties can be set on saveOptions here.
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine("File has been decrypted and saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
