using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Encryption Test");
            sheet.Cells["A2"].PutValue(DateTime.Now);
            sheet.Cells["A3"].PutValue(12345);

            // Define file paths for the unencrypted and encrypted workbooks
            string unencryptedPath = "UnencryptedWorkbook.xlsx";
            string encryptedPath = "EncryptedWorkbook.xlsx";

            // Save the workbook without any protection
            workbook.Save(unencryptedPath, SaveFormat.Xlsx);

            // Get the file size of the unencrypted workbook
            long unencryptedSize = new FileInfo(unencryptedPath).Length;
            Console.WriteLine($"Unencrypted file size: {unencryptedSize} bytes");

            // Apply strong encryption settings
            // Set a password that will be required to open the workbook
            workbook.Settings.Password = "StrongPassword123";

            // Use the StrongCryptographicProvider encryption type with a 128‑bit key
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the encrypted workbook
            workbook.Save(encryptedPath, SaveFormat.Xlsx);

            // Get the file size of the encrypted workbook
            long encryptedSize = new FileInfo(encryptedPath).Length;
            Console.WriteLine($"Encrypted file size: {encryptedSize} bytes");

            // Compare the sizes
            if (encryptedSize > unencryptedSize)
            {
                Console.WriteLine("Encryption increased the file size.");
            }
            else if (encryptedSize < unencryptedSize)
            {
                Console.WriteLine("Encryption decreased the file size.");
            }
            else
            {
                Console.WriteLine("File size unchanged after encryption.");
            }

            // Clean up
            workbook.Dispose();
        }
    }
}