using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    class Program
    {
        static void Main()
        {
            // Paths for the test files
            string unencryptedPath = "UnencryptedWorkbook.xlsx";
            string encryptedPath = "EncryptedWorkbook_AES256.xlsx";

            // -----------------------------------------------------------------
            // 1. Create a new workbook and add some sample data
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Encryption Test");
            sheet.Cells["A2"].PutValue(DateTime.Now);
            sheet.Cells["A3"].PutValue(12345);

            // -----------------------------------------------------------------
            // 2. Save the workbook without any protection (baseline size)
            // -----------------------------------------------------------------
            workbook.Save(unencryptedPath, SaveFormat.Xlsx);
            long unencryptedSize = new FileInfo(unencryptedPath).Length;
            Console.WriteLine($"Unencrypted file size: {unencryptedSize} bytes");

            // -----------------------------------------------------------------
            // 3. Apply password protection and AES‑256 encryption options
            // -----------------------------------------------------------------
            workbook.Settings.Password = "StrongPassword123!";
            // EncryptionType is ignored for XLSX, but required by the API
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

            // -----------------------------------------------------------------
            // 4. Save the encrypted workbook
            // -----------------------------------------------------------------
            workbook.Save(encryptedPath, SaveFormat.Xlsx);
            long encryptedSize = new FileInfo(encryptedPath).Length;
            Console.WriteLine($"Encrypted file size: {encryptedSize} bytes");

            // -----------------------------------------------------------------
            // 5. Verify that encryption increased the file size
            // -----------------------------------------------------------------
            if (encryptedSize > unencryptedSize)
                Console.WriteLine("File size increased after encryption as expected.");
            else
                Console.WriteLine("Warning: Encrypted file size is not larger than the unencrypted file.");

            // -----------------------------------------------------------------
            // 6. Load the encrypted workbook using the password to confirm it works
            // -----------------------------------------------------------------
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = "StrongPassword123!";
            Workbook loadedEncrypted = new Workbook(encryptedPath, loadOptions);
            Console.WriteLine("Encrypted workbook loaded successfully.");
            Console.WriteLine($"Cell A1 value after load: {loadedEncrypted.Worksheets[0].Cells["A1"].StringValue}");

            // Cleanup
            workbook.Dispose();
            loadedEncrypted.Dispose();
        }
    }
}