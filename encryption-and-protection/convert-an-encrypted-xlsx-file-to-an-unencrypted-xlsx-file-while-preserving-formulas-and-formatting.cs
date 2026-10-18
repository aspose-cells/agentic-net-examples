// Title: Decrypt a password‑protected XLSX file and save it as an unencrypted workbook while preserving formulas and formatting using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells load password protected XLSX and save without password | How to open an encrypted Excel file with Aspose.Cells and keep all formulas | Remove workbook encryption in .NET while preserving cell formatting using Aspose.Cells | Decrypt an XLSX file programmatically with Aspose.Cells LoadOptions | Save a protected Excel workbook as plain XLSX in C#
// Tags: Aspose.Cells LoadOptions password | decrypt password protected XLSX Aspose.Cells | save unencrypted XLSX Aspose.Cells | preserve formulas during workbook decryption | C# remove Excel file encryption programmatically

using System;
using System.IO;
using Aspose.Cells;

// The example loads a password‑protected XLSX workbook using Aspose.Cells LoadOptions with the supplied password, then saves it to a new XLSX file without a password, ensuring that all formulas, formatting, and other content remain unchanged.
class Program
{
    static void Main()
    {
        // Path to the encrypted XLSX file
        string encryptedFilePath = @"C:\Path\To\EncryptedFile.xlsx";

        // Password used to encrypt the file
        string password = "YourPassword";

        // Path where the unencrypted XLSX will be saved
        string unencryptedFilePath = @"C:\Path\To\UnencryptedFile.xlsx";

        try
        {
            // Verify that the encrypted file exists before attempting to load it
            if (!File.Exists(encryptedFilePath))
                throw new FileNotFoundException("Encrypted file not found.", encryptedFilePath);

            // LoadOptions allows specifying the password for opening the encrypted workbook
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };

            // Load the encrypted workbook; formulas, formatting, and other content are preserved
            Workbook workbook = new Workbook(encryptedFilePath, loadOptions);

            // Save the workbook without specifying a password – it will be saved as an unencrypted file
            workbook.Save(unencryptedFilePath, SaveFormat.Xlsx);

            Console.WriteLine("Conversion completed successfully.");
        }
        catch (FileNotFoundException ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
