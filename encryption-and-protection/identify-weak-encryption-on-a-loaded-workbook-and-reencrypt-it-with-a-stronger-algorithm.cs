using System;
using Aspose.Cells;

class ReEncryptWorkbook
{
    static void Main()
    {
        // Paths for the source and destination workbooks
        string sourcePath = "input.xlsx";
        string destinationPath = "output_strong_encrypted.xlsx";

        // Password to open the existing workbook (if it is encrypted)
        // Replace with the actual password or leave empty if not needed
        string existingPassword = "oldPassword";

        // Password to apply for the stronger encryption
        string newPassword = "newStrongPassword";

        // Load the workbook with the existing password (if any)
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = existingPassword;
        Workbook workbook = new Workbook(sourcePath, loadOptions);

        // Determine if the workbook is currently encrypted
        bool isEncrypted = workbook.Settings.IsEncrypted;

        // Apply strong encryption (AES 256) regardless of current state
        // Set the new password
        workbook.Settings.Password = newPassword;

        // Use the strongest available encryption type with a 256‑bit key
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

        // Save the workbook with the new encryption settings
        workbook.Save(destinationPath, SaveFormat.Xlsx);
    }
}