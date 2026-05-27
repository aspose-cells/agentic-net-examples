using System;
using Aspose.Cells;

class ApplyDefaultEncryption
{
    static void Main()
    {
        // Input and output file paths
        string inputPath = "input.xlsx";
        string outputPath = "output_encrypted.xlsx";

        // Preset password to be applied if the workbook is not encrypted
        string presetPassword = "MySecret123";

        // Load the workbook (no password needed for unencrypted files)
        LoadOptions loadOptions = new LoadOptions();
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Check if the workbook is already encrypted
        if (!workbook.Settings.IsEncrypted)
        {
            // Apply password protection
            workbook.Settings.Password = presetPassword;

            // Optionally set stronger encryption options (e.g., AES 128-bit)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);
        }

        // Save the workbook with the applied encryption
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}