using System;
using System.IO;
using Aspose.Cells;

namespace EncryptionImpactDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample data for encryption impact analysis.");
            sheet.Cells["A2"].PutValue(DateTime.Now);
            sheet.Cells["A3"].PutValue(12345);

            // Define file paths for unencrypted and encrypted workbooks
            string unencryptedPath = "UnencryptedWorkbook.xlsx";
            string encryptedPath = "EncryptedWorkbook.xlsx";

            // Save the unencrypted workbook
            workbook.Save(unencryptedPath, SaveFormat.Xlsx);

            // Get file size of the unencrypted workbook
            long unencryptedSize = new FileInfo(unencryptedPath).Length;

            // Apply password protection (encryption) to the workbook
            workbook.Settings.Password = "StrongPassword123";
            // Optional: set encryption options (e.g., AES 128-bit)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the encrypted workbook
            workbook.Save(encryptedPath, SaveFormat.Xlsx);

            // Get file size of the encrypted workbook
            long encryptedSize = new FileInfo(encryptedPath).Length;

            // Output the sizes and the size difference
            Console.WriteLine($"Unencrypted file size: {unencryptedSize} bytes");
            Console.WriteLine($"Encrypted file size:   {encryptedSize} bytes");
            Console.WriteLine($"Size increase:         {encryptedSize - unencryptedSize} bytes");

            // Verify encryption status using FileFormatInfo
            FileFormatInfo unencryptedInfo = FileFormatUtil.DetectFileFormat(unencryptedPath);
            FileFormatInfo encryptedInfo = FileFormatUtil.DetectFileFormat(encryptedPath);
            Console.WriteLine($"Is unencrypted file encrypted? {unencryptedInfo.IsEncrypted}");
            Console.WriteLine($"Is encrypted file encrypted?   {encryptedInfo.IsEncrypted}");

            // Clean up
            workbook.Dispose();
        }
    }
}