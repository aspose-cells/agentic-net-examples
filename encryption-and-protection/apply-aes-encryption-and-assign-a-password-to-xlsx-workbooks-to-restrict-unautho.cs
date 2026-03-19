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

            // Add some data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive data");

            // Set the password that will be required to open the workbook
            workbook.Settings.Password = "StrongPassword123";

            // Apply AES encryption (StrongCryptographicProvider) with a 256‑bit key
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

            // Save the encrypted workbook as .xlsx
            string encryptedPath = "EncryptedWorkbook.xlsx";
            workbook.Save(encryptedPath, SaveFormat.Xlsx);

            // ---- Verify by loading the protected workbook ----
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = "StrongPassword123";

            Workbook loadedWorkbook = new Workbook(encryptedPath, loadOptions);
            string cellValue = loadedWorkbook.Worksheets[0].Cells["A1"].StringValue;

            Console.WriteLine("Loaded cell value: " + cellValue);
        }
    }
}