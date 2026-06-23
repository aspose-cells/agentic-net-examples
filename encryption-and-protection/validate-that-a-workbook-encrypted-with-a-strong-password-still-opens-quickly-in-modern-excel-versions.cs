using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionValidation
{
    class Program
    {
        static void Main()
        {
            // Path for the encrypted workbook
            string encryptedFilePath = "EncryptedStrong.xlsx";

            // -----------------------------------------------------------------
            // 1. Create a new workbook and add some sample data
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Encryption Test");
            sheet.Cells["A2"].PutValue(DateTime.Now);
            sheet.Cells["A3"].PutValue(12345);

            // -----------------------------------------------------------------
            // 2. Apply a strong password and strong encryption options
            // -----------------------------------------------------------------
            workbook.Settings.Password = "StrongPassword123!"; // set password
            // Use strong encryption (AES 256) – EncryptionType.StrongCryptographicProvider
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

            // -----------------------------------------------------------------
            // 3. Save the encrypted workbook
            // -----------------------------------------------------------------
            workbook.Save(encryptedFilePath, SaveFormat.Xlsx);

            // -----------------------------------------------------------------
            // 4. Verify that the file is reported as encrypted
            // -----------------------------------------------------------------
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(encryptedFilePath);
            Console.WriteLine($"File detected as encrypted: {formatInfo.IsEncrypted}");

            // -----------------------------------------------------------------
            // 5. Measure the time required to open the encrypted workbook in modern Excel (simulated by Aspose.Cells)
            // -----------------------------------------------------------------
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = "StrongPassword123!"
            };

            Stopwatch sw = Stopwatch.StartNew();
            Workbook loadedWorkbook = new Workbook(encryptedFilePath, loadOptions);
            sw.Stop();

            Console.WriteLine($"Time to open encrypted workbook: {sw.ElapsedMilliseconds} ms");

            // -----------------------------------------------------------------
            // 6. Confirm that the loaded workbook reports encryption status correctly
            // -----------------------------------------------------------------
            Console.WriteLine($"Loaded workbook IsEncrypted: {loadedWorkbook.Settings.IsEncrypted}");

            // Cleanup (optional)
            workbook.Dispose();
            loadedWorkbook.Dispose();
        }
    }
}