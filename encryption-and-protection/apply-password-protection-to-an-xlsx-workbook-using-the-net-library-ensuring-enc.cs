using System;
using Aspose.Cells;
using Aspose.Cells.Metadata;

namespace AsposeCellsPasswordProtectionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Add some sample data to the first worksheet
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive Data");

            // Set a password to encrypt the workbook file
            wb.Settings.Password = "StrongPassword123";

            // (Optional) Define stronger encryption options
            wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the encrypted workbook
            string encryptedPath = "EncryptedWorkbook.xlsx";
            wb.Save(encryptedPath, SaveFormat.Xlsx);

            // Load the encrypted workbook using the password
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = "StrongPassword123";
            Workbook loadedWb = new Workbook(encryptedPath, loadOptions);

            // Verify that the data can be accessed after decryption
            string cellValue = loadedWb.Worksheets[0].Cells["A1"].StringValue;
            Console.WriteLine("Decrypted cell value: " + cellValue);
        }
    }
}