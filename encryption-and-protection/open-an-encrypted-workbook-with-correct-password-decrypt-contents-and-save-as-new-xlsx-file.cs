using System;
using System.IO;
using Aspose.Cells;

class DecryptWorkbook
{
    static void Main()
    {
        // Paths and password
        string encryptedFile = "encrypted.xlsx";
        string password = "yourPassword";
        string decryptedFile = "decrypted.xlsx";

        // Verify that the encrypted file exists
        if (!File.Exists(encryptedFile))
        {
            Console.WriteLine($"Error: File \"{encryptedFile}\" not found.");
            return;
        }

        try
        {
            // Load the encrypted workbook with the password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };
            Workbook workbook = new Workbook(encryptedFile, loadOptions);

            // Remove the password protection
            workbook.Settings.Password = null;

            // Save the workbook as a new unencrypted XLSX file
            workbook.Save(decryptedFile, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook decrypted and saved to \"{decryptedFile}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}