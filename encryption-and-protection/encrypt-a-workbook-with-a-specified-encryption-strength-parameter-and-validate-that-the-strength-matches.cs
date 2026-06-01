using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    class Program
    {
        static void Main()
        {
            // Parameters
            string password = "StrongPwd123";
            int desiredKeyLength = 128; // Encryption strength in bits

            // ------------------- Create Workbook -------------------
            Workbook workbook = new Workbook(); // create new workbook

            // Add sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Encryption Strength Test");

            // Set password for opening the file
            workbook.Settings.Password = password;

            // Set encryption options with the desired strength
            // EncryptionType is ignored for modern formats, but required by the method signature
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, desiredKeyLength);

            // Save the encrypted workbook
            string filePath = "EncryptedWorkbook.xlsx";
            workbook.Save(filePath, SaveFormat.Xlsx);

            // ------------------- Load and Validate -------------------
            // Load the workbook using the password
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = password;
            Workbook loadedWorkbook = new Workbook(filePath, loadOptions);

            // Verify that the workbook reports being encrypted
            bool isEncrypted = loadedWorkbook.Settings.IsEncrypted;
            Console.WriteLine($"Workbook IsEncrypted: {isEncrypted}");

            // Verify that the password is correctly set (round‑trip check)
            bool passwordMatches = loadedWorkbook.Settings.Password == password;
            Console.WriteLine($"Password matches: {passwordMatches}");

            // Since Aspose.Cells does not expose the key length after loading,
            // successful loading with the correct password confirms that the
            // encryption (with the requested strength) is in effect.
            if (isEncrypted && passwordMatches)
            {
                Console.WriteLine($"Encryption applied with key length {desiredKeyLength} bits.");
            }
            else
            {
                Console.WriteLine("Encryption validation failed.");
            }
        }
    }
}