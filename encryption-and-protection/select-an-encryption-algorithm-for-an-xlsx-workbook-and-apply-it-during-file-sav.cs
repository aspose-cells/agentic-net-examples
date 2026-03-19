using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default empty workbook)
            Workbook workbook = new Workbook();

            // Add some sample data (optional, just to have content)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive Data");
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // Set a password that will be required to open the workbook
            workbook.Settings.Password = "MyStrongPassword";

            // Choose an encryption algorithm and key length.
            // StrongCryptographicProvider corresponds to AES encryption used by modern Excel versions.
            // Key length can be 40, 128, or 256 bits. Here we use 256 for maximum security.
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

            // Save the workbook as an encrypted XLSX file
            workbook.Save("EncryptedWorkbook.xlsx", SaveFormat.Xlsx);

            // Optional: Verify that the workbook is encrypted by loading it with the password
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = "MyStrongPassword";
            Workbook loadedWorkbook = new Workbook("EncryptedWorkbook.xlsx", loadOptions);

            // Output a confirmation
            Console.WriteLine("Workbook saved with encryption and successfully loaded.");
        }
    }
}