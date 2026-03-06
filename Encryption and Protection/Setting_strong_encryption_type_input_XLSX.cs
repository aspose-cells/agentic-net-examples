using System;
using Aspose.Cells;

class SetStrongEncryption
{
    static void Main()
    {
        // Load the existing XLSX workbook
        string inputPath = "input.xlsx";
        LoadOptions loadOptions = new LoadOptions(); // no password needed for unencrypted file
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Set a password that will be required to open the workbook
        workbook.Settings.Password = "StrongPassword123";

        // Apply strong encryption (AES 128-bit) to the workbook
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Optionally protect the workbook structure with the same password
        workbook.Protect(ProtectionType.Structure, "StrongPassword123");

        // Save the encrypted workbook
        string outputPath = "encrypted_output.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}