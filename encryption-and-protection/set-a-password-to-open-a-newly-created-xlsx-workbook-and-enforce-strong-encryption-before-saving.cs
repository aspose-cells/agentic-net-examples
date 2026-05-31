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
            sheet.Cells["A1"].PutValue("Sensitive information");

            // Set the password required to open the workbook
            workbook.Settings.Password = "StrongPass!123";

            // Enforce strong encryption (AES 256-bit)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

            // Save the encrypted workbook as XLSX
            workbook.Save("EncryptedWorkbook.xlsx", SaveFormat.Xlsx);

            // Optional: Verify that the workbook can be opened with the password
            LoadOptions loadOptions = new LoadOptions { Password = "StrongPass!123" };
            Workbook loaded = new Workbook("EncryptedWorkbook.xlsx", loadOptions);
            Console.WriteLine("Loaded cell value: " + loaded.Worksheets[0].Cells["A1"].Value);
        }
    }
}