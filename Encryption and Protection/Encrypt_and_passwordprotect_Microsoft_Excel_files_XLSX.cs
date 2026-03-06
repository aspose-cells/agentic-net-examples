using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive Data");
            sheet.Cells["B2"].PutValue(DateTime.Now);

            // Set the workbook password (this encrypts the file)
            workbook.Settings.Password = "MySecretPassword";

            // Optional: specify encryption algorithm and key length (ignored for .xlsx but kept for completeness)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Protect the workbook structure with the same password
            workbook.Protect(ProtectionType.Structure, "MySecretPassword");

            // Save the encrypted and protected workbook
            string outputPath = "EncryptedProtectedWorkbook.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            // Demonstrate loading the protected workbook using the password
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = "MySecretPassword";
            Workbook loadedWorkbook = new Workbook(outputPath, loadOptions);

            // Verify that the data is accessible after decryption
            string cellValue = loadedWorkbook.Worksheets[0].Cells["A1"].StringValue;
            Console.WriteLine("Loaded cell A1 value: " + cellValue);
        }
    }
}