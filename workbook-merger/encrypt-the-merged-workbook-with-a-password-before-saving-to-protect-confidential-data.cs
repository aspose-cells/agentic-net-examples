using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing merged workbook)
            Workbook wb = new Workbook();

            // Example: add some data to the first worksheet
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Confidential Data");

            // Set the password that will be required to open the workbook
            wb.Settings.Password = "StrongPassword123";

            // (Optional) Define encryption strength – 128‑bit AES is default for modern formats
            // wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the encrypted workbook
            wb.Save("MergedWorkbook_Encrypted.xlsx");

            // Verify that the workbook is encrypted by loading it with the password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = "StrongPassword123"
            };
            Workbook loadedWb = new Workbook("MergedWorkbook_Encrypted.xlsx", loadOptions);
            Console.WriteLine("Loaded cell value: " + loadedWb.Worksheets[0].Cells["A1"].Value);
        }
    }
}