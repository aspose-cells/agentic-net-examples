using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source XLSX file
        string inputPath = "input.xlsx";

        // Load the workbook (no password needed for an unencrypted file)
        LoadOptions loadOptions = new LoadOptions();
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Set a password that will protect the workbook
        workbook.Settings.Password = "myPassword";

        // Choose the encryption algorithm and key length.
        // Example: StrongCryptographicProvider with a 256‑bit key.
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

        // Save the encrypted workbook
        string outputPath = "encrypted.xlsx";
        workbook.Save(outputPath);
    }
}