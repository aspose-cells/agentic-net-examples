using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsEncryptionTiming
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Encryption test");
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // Set a password to encrypt the workbook
            workbook.Settings.Password = "SecretPassword123";

            // Optional: specify encryption algorithm and key length (strong encryption)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the encrypted workbook to disk
            string encryptedFilePath = "encrypted_workbook.xlsx";
            workbook.Save(encryptedFilePath);
            workbook.Dispose();

            // Measure the time required to open the encrypted workbook in a headless environment
            Stopwatch stopwatch = Stopwatch.StartNew();

            // LoadOptions with the password are required to open the encrypted file
            LoadOptions loadOptions = new LoadOptions { Password = "SecretPassword123" };
            Workbook loadedWorkbook = new Workbook(encryptedFilePath, loadOptions);

            stopwatch.Stop();

            Console.WriteLine($"Time to open encrypted workbook: {stopwatch.ElapsedMilliseconds} ms");

            // Verify that the workbook was loaded correctly (optional)
            Console.WriteLine($"Cell A1 value: {loadedWorkbook.Worksheets[0].Cells["A1"].StringValue}");

            loadedWorkbook.Dispose();
        }
    }
}