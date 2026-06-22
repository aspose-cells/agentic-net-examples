using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class DecryptWorkbook
    {
        public static void Run()
        {
            // Path to the encrypted XLSX file
            string encryptedPath = "encrypted.xlsx";

            // Path for the unencrypted output file
            string decryptedPath = "decrypted.xlsx";

            // Password used to open the encrypted workbook
            string password = "yourPassword";

            // Verify the encrypted file exists
            if (!File.Exists(encryptedPath))
            {
                Console.WriteLine($"Error: Encrypted file not found at '{encryptedPath}'.");
                return;
            }

            try
            {
                // Load the encrypted workbook using LoadOptions with the password
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto) { Password = password };
                Workbook workbook = new Workbook(encryptedPath, loadOptions);

                // Ensure no password is set for saving (remove any workbook-level protection if present)
                workbook.Settings.Password = null;

                // Save the workbook as an unencrypted XLSX file, preserving formulas and formatting
                workbook.Save(decryptedPath, SaveFormat.Xlsx);

                Console.WriteLine($"Decryption completed. Unencrypted file saved to: {decryptedPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            DecryptWorkbook.Run();
        }
    }
}