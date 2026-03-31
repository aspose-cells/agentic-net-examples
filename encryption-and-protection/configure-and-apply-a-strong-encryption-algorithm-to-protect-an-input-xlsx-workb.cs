using System;
using Aspose.Cells;

class EncryptWorkbook
{
    static void Main()
    {
        // Input XLSX file to be encrypted
        string inputPath = "input.xlsx";

        // Output path for the encrypted workbook
        string outputPath = "encrypted.xlsx";

        // Load the existing workbook
        Workbook workbook = new Workbook(inputPath);

        // Set the password required to open the workbook
        workbook.Settings.Password = "StrongPassword!123";

        // Apply strong encryption (AES with 256‑bit key)
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

        // Save the workbook; the encryption settings are persisted automatically
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}