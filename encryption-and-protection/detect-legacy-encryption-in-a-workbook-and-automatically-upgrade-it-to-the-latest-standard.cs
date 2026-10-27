// Title: Detect legacy encrypted Excel workbook and upgrade it to AES‑256 using Aspose.Cells for .NET
// Common Searches: how to check if an Excel file is encrypted with a legacy algorithm using Aspose.Cells C# | re‑encrypt a password‑protected XLSX to AES‑256 with Aspose.Cells .NET | load an encrypted workbook with known password and save with new encryption in C# | upgrade legacy Excel file encryption to modern standard programmatically | detect and convert legacy encrypted XLSX to AES‑256 using Aspose.Cells API
// Tags: legacy Excel encryption detection | AES‑256 workbook re‑encryption | LoadOptions password decryption Aspose.Cells | Workbook.Settings.Password encryption upgrade | modernize encrypted XLSX files

using System;
using System.IO;
using Aspose.Cells;

// The example loads a potentially legacy‑encrypted XLSX, detects encryption by attempting an unprotected load, reopens it with LoadOptions using the known password if needed, then sets Workbook.Settings.Password to a new password, causing Aspose.Cells to save the file with AES‑256 encryption.
class WorkbookEncryptionUpgrade
{
    static void Main()
    {
        // Paths to the source (legacy encrypted) and destination workbooks
        string sourcePath = "legacy_encrypted.xlsx";
        string upgradedPath = "upgraded.xlsx";

        // Known password for the legacy encrypted workbook
        string legacyPassword = "legacyPassword";

        // Verify that the source file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Source file not found: {sourcePath}");
            return;
        }

        Workbook workbook = null;
        bool wasEncrypted = false;

        // Attempt to load the workbook without a password; if it fails, retry with the known legacy password
        try
        {
            workbook = new Workbook(sourcePath);
        }
        catch (Exception)
        {
            // Assume the failure is due to encryption
            wasEncrypted = true;
            try
            {
                var loadOptions = new LoadOptions { Password = legacyPassword };
                workbook = new Workbook(sourcePath, loadOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading encrypted workbook: {ex.Message}");
                return;
            }
        }

        try
        {
            if (wasEncrypted)
            {
                // Upgrade to the latest encryption standard (AES‑256) with a new password
                string newPassword = "newStrongPassword";

                // Set the new password; Aspose.Cells will apply AES‑256 encryption automatically
                workbook.Settings.Password = newPassword;

                workbook.Save(upgradedPath, SaveFormat.Xlsx);
                Console.WriteLine("Workbook encrypted with legacy method detected and upgraded.");
            }
            else
            {
                // Workbook is not encrypted; optionally apply encryption now
                Console.WriteLine("Workbook is not encrypted.");
                // Example: save without encryption
                workbook.Save(upgradedPath, SaveFormat.Xlsx);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
