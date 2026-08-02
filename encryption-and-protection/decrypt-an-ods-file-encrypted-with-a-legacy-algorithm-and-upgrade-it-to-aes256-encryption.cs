// Title: C# – Decrypt a legacy‑encrypted ODS file and re‑encrypt it with AES‑256 using Aspose.Cells
// Description: Step‑by‑step sample that detects encryption on an ODS workbook, opens it with the old password via LoadOptions, assigns a new strong password, applies AES‑256 protection with SetEncryptionOptions, and saves the file with OdsSaveOptions while handling missing files and runtime errors.
// Keywords: Aspose.Cells ODS decryption | AES‑256 encryption Aspose.Cells | legacy ODS password C# | upgrade ODS protection .NET | SetEncryptionOptions example | LoadOptions password ODS | secure ODS files
// Common Searches: how to change password of an encrypted ODS using Aspose.Cells | convert legacy ODS encryption to AES‑256 in C# | detect encrypted ODS workbook with Aspose.Cells | re‑encrypt ODS file with stronger algorithm .NET | Aspose.Cells example for ODS password upgrade
// Developer Intent: The developer needs to open an ODS spreadsheet protected by an old algorithm, remove the legacy password, and save it again secured with AES‑256.
// Use Cases: Modernize archived ODS spreadsheets to comply with current security policies. | Batch‑process a repository of ODS files, replacing weak passwords with strong AES‑256 encryption. | Validate encryption status before integrating ODS files into an automated data‑pipeline.
// AI Prompts: Write C# code that loads an ODS file encrypted with a legacy password, then saves it using AES‑256 encryption with Aspose.Cells. | Explain how to programmatically check if an ODS workbook is encrypted before opening it in Aspose.Cells. | Suggest robust error‑handling patterns for re‑encrypting ODS files with Aspose.Cells in a production environment.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Ods;

// Step‑by‑step sample that detects encryption on an ODS workbook, opens it with the old password via LoadOptions, assigns a new strong password, applies AES‑256 protection with SetEncryptionOptions, and saves the file with OdsSaveOptions while handling missing files and runtime errors.
class OdsEncryptionUpgrade
{
    static void Main()
    {
        // Paths to the source (legacy encrypted) and destination (AES‑256 encrypted) ODS files
        string inputPath = "legacy_encrypted.ods";
        string outputPath = "upgraded_encrypted.ods";

        // Password used by the legacy encryption algorithm
        string legacyPassword = "oldPassword";

        // New password for AES‑256 encryption
        string newPassword = "newStrongPassword";

        try
        {
            // Verify that the source file exists before attempting to process it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Source file not found: {inputPath}");
                return;
            }

            // -----------------------------------------------------------------
            // 1. Detect whether the source file is encrypted
            // -----------------------------------------------------------------
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(inputPath);
            Console.WriteLine($"Is the source file encrypted? {formatInfo.IsEncrypted}");

            // -----------------------------------------------------------------
            // 2. Load the ODS file using the legacy password
            // -----------------------------------------------------------------
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Ods)
            {
                Password = legacyPassword // decrypt with legacy password
            };
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // -----------------------------------------------------------------
            // 3. Apply new AES‑256 encryption
            // -----------------------------------------------------------------
            workbook.Settings.Password = newPassword; // set new password
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256); // AES‑256

            // -----------------------------------------------------------------
            // 4. Save the workbook with ODS save options (default ODS format)
            // -----------------------------------------------------------------
            OdsSaveOptions saveOptions = new OdsSaveOptions(); // default options
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine("File has been re‑encrypted with AES‑256 and saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
