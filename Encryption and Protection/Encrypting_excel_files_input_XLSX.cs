using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file
            string sourceFile = "input.xlsx";

            // Path for the encrypted output file
            string encryptedFile = "encrypted.xlsx";

            // Password to protect the workbook
            string password = "StrongPassword123";

            // Load the existing workbook
            Workbook workbook = new Workbook(sourceFile);

            // Set the password for opening the workbook
            workbook.Settings.Password = password;

            // Define encryption type and key length (e.g., StrongCryptographicProvider with 128‑bit key)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the workbook; the password and encryption settings are applied automatically
            workbook.Save(encryptedFile, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook encrypted and saved to '{encryptedFile}'.");
        }
    }
}