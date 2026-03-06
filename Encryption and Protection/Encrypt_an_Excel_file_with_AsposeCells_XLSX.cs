using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive Data");

            // Set a password to encrypt the workbook (XLSX uses SHA‑AES encryption automatically)
            workbook.Settings.Password = "MySecretPassword";

            // Optional: specify encryption options (ignored for XLSX but shown for completeness)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the encrypted workbook (lifecycle: save)
            string encryptedPath = "EncryptedWorkbook.xlsx";
            workbook.Save(encryptedPath, SaveFormat.Xlsx);

            // Verify that the file is encrypted by loading it with the password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = "MySecretPassword"
            };
            Workbook loadedWorkbook = new Workbook(encryptedPath, loadOptions);
            Console.WriteLine("Loaded encrypted workbook successfully. Cell A1 value: " +
                              loadedWorkbook.Worksheets[0].Cells["A1"].Value);
        }
    }
}