using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionSizeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some data to make the file sizable
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            for (int row = 0; row < 1000; row++)
            {
                for (int col = 0; col < 20; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row}C{col}");
                }
            }

            // Save the workbook without encryption
            string unencryptedPath = "UnencryptedWorkbook.xlsx";
            workbook.Save(unencryptedPath);
            long unencryptedSize = new FileInfo(unencryptedPath).Length;

            // Apply password protection (encryption)
            workbook.Settings.Password = "StrongPassword123";
            // Optionally set stronger encryption options (e.g., 128-bit AES)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the encrypted workbook
            string encryptedPath = "EncryptedWorkbook.xlsx";
            workbook.Save(encryptedPath);
            long encryptedSize = new FileInfo(encryptedPath).Length;

            // Output the file sizes for comparison
            Console.WriteLine($"Unencrypted file size: {unencryptedSize} bytes");
            Console.WriteLine($"Encrypted file size:   {encryptedSize} bytes");
            Console.WriteLine($"Size increase:        {encryptedSize - unencryptedSize} bytes");

            // Clean up
            workbook.Dispose();
        }
    }
}