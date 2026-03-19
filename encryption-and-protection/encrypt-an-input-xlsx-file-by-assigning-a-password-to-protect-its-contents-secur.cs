using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source XLSX file (must exist)
            string sourcePath = "input.xlsx";

            // Path where the encrypted file will be saved
            string encryptedPath = "input_encrypted.xlsx";

            // Password to protect the workbook
            string password = "StrongPassword123";

            // Load the existing workbook (no password needed for unprotected file)
            Workbook workbook = new Workbook(sourcePath);

            // Assign a password to encrypt the workbook
            workbook.Settings.Password = password;

            // Save the workbook; the file will be encrypted with the specified password
            workbook.Save(encryptedPath);

            // Optional: Verify that the workbook is now encrypted
            Workbook encryptedWorkbook = new Workbook(encryptedPath, new LoadOptions { Password = password });
            Console.WriteLine("Encryption successful. Cell A1 value: " + encryptedWorkbook.Worksheets[0].Cells["A1"].Value);
        }
    }
}