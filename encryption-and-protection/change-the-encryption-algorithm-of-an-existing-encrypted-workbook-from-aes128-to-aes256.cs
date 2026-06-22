using Aspose.Cells;
using System;
using System.IO;

class ChangeEncryption
{
    static void Main()
    {
        try
        {
            // Path to the existing AES‑128 encrypted workbook
            string inputPath = "Encrypted128.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Password that protects the current workbook
            string password = "myPassword";

            // Load the workbook with the existing password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = password
            };
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Change encryption to AES‑256 (StrongCryptographicProvider with 256‑bit key)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

            // Re‑apply the password so the workbook remains protected when saved
            workbook.Settings.Password = password;

            // Save the workbook with the new encryption settings
            string outputPath = "Encrypted256.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}