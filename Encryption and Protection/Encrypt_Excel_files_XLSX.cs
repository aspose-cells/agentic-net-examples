using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive data");

            // Set a password to encrypt the workbook (WorkbookSettings.Password property)
            workbook.Settings.Password = "MySecretPassword";

            // Optionally set encryption options (ignored for .xlsx but shown for completeness)
            // workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the encrypted workbook (lifecycle rule: save)
            string encryptedPath = "EncryptedWorkbook.xlsx";
            workbook.Save(encryptedPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved with encryption to '{encryptedPath}'.");

            // Load the encrypted workbook using the password (lifecycle rule: load)
            LoadOptions loadOptions = new LoadOptions
            {
                Password = "MySecretPassword"
            };
            Workbook loadedWorkbook = new Workbook(encryptedPath, loadOptions);

            // Verify that the data can be read
            string cellValue = loadedWorkbook.Worksheets[0].Cells["A1"].StringValue;
            Console.WriteLine($"Loaded cell A1 value: {cellValue}");

            // Clean up
            workbook.Dispose();
            loadedWorkbook.Dispose();
        }
    }
}