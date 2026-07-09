using System;
using Aspose.Cells;

class ApplyDefaultEncryption
{
    static void Main()
    {
        // Input and output file paths
        string inputPath = "input.xlsx";
        string outputPath = "output_encrypted.xlsx";

        // Preset password to be applied when encrypting
        const string presetPassword = "MySecretPassword";

        // Load the workbook (no password needed for unencrypted files)
        LoadOptions loadOptions = new LoadOptions();
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Check if the workbook is already encrypted
        if (!workbook.Settings.IsEncrypted)
        {
            // Apply password protection
            workbook.Settings.Password = presetPassword;

            // Enable default encryption (encrypt with default password when structure/windows are locked)
            workbook.Settings.IsDefaultEncrypted = true;

            // Optionally set stronger encryption options (e.g., AES 128-bit)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);
        }

        // Save the workbook with encryption applied
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}