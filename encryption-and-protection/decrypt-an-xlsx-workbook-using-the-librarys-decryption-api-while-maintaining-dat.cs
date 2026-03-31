using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDecryptionDemo
{
    class Program
    {
        static void Main()
        {
            string encryptedFilePath = "encrypted_workbook.xlsx";
            string password = "mySecretPassword";

            // Create a sample encrypted workbook if it does not exist
            if (!File.Exists(encryptedFilePath))
            {
                Workbook sample = new Workbook();
                sample.Worksheets[0].Cells["A1"].PutValue("Secret Data");
                sample.Settings.Password = password; // encrypt
                sample.Save(encryptedFilePath);
                Console.WriteLine($"Sample encrypted workbook created at '{encryptedFilePath}'.");
            }

            // Load the encrypted workbook with the password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx) { Password = password };
            Workbook workbook = new Workbook(encryptedFilePath, loadOptions);

            Console.WriteLine("Is workbook encrypted after load? " + workbook.Settings.IsEncrypted);

            // Remove encryption
            workbook.Settings.Password = null;

            // Save the decrypted workbook
            string decryptedFilePath = "decrypted_workbook.xlsx";
            workbook.Save(decryptedFilePath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook decrypted and saved to '{decryptedFilePath}'.");
        }
    }
}