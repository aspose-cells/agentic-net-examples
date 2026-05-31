using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionValidation
{
    class Program
    {
        static void Main()
        {
            // Path for the encrypted workbook
            string encryptedFilePath = "encrypted.xlsx";

            // 1. Create a new workbook and add some data
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Encryption test");

            // 2. Set a password to encrypt the workbook
            wb.Settings.Password = "SecretPassword";

            // 3. Save the workbook (encryption is applied automatically)
            wb.Save(encryptedFilePath);

            // 4. Verify encryption status using FileFormatInfo
            FileFormatInfo fileInfo = FileFormatUtil.DetectFileFormat(encryptedFilePath);
            Console.WriteLine($"FileFormatInfo.IsEncrypted: {fileInfo.IsEncrypted}");

            // 5. Verify encryption status using WorkbookSettings after loading with password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto)
            {
                Password = "SecretPassword"
            };
            Workbook loadedWb = new Workbook(encryptedFilePath, loadOptions);
            Console.WriteLine($"Workbook.Settings.IsEncrypted after load: {loadedWb.Settings.IsEncrypted}");

            // 6. Attempt to load without password to demonstrate failure (optional)
            try
            {
                Workbook failLoad = new Workbook(encryptedFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Loading without password failed as expected: {ex.Message}");
            }
        }
    }
}