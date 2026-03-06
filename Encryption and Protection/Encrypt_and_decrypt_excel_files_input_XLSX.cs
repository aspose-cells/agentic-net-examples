using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    class Program
    {
        static void Main()
        {
            // Path for the original and encrypted files
            string originalPath = "OriginalWorkbook.xlsx";
            string encryptedPath = "EncryptedWorkbook.xlsx";

            // ------------------- Create and encrypt workbook -------------------
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive Data");
            sheet.Cells["B2"].PutValue(DateTime.Now);

            // Set a password to protect the workbook
            workbook.Settings.Password = "MySecretPassword";

            // Optional: set encryption options (ignored for .xlsx but shown for completeness)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the workbook; it will be encrypted with the password
            workbook.Save(encryptedPath, SaveFormat.Xlsx);

            // ------------------- Verify encryption -------------------
            // Detect if the saved file is encrypted
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(encryptedPath);
            Console.WriteLine($"Is the file encrypted? {formatInfo.IsEncrypted}");

            // ------------------- Load (decrypt) workbook -------------------
            // Prepare load options with the password
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = "MySecretPassword";

            // Load the encrypted workbook using the password
            Workbook loadedWorkbook = new Workbook(encryptedPath, loadOptions);

            // Access data to confirm successful decryption
            string cellValue = loadedWorkbook.Worksheets[0].Cells["A1"].StringValue;
            Console.WriteLine($"Decrypted cell A1 value: {cellValue}");

            // Optionally, save the decrypted workbook to a new file
            loadedWorkbook.Save(originalPath, SaveFormat.Xlsx);
        }
    }
}