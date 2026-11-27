// Title: Decrypt a password‑protected Excel workbook and verify cell data using Aspose.Cells for .NET
// Common Searches: C# how to open a password protected .xlsx file with Aspose.Cells LoadOptions | verify data integrity after decrypting an Excel workbook using Aspose.Cells | remove password from an encrypted workbook programmatically with Aspose.Cells C# | read specific cell value from a decrypted workbook using Aspose.Cells | handle FileNotFoundException when loading an encrypted Excel file with Aspose.Cells
// Tags: Aspose.Cells LoadOptions password decryption | C# open encrypted xlsx Aspose.Cells | Aspose.Cells read cell after decryption | Aspose.Cells save workbook without password | Aspose.Cells encrypted workbook exception handling

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the presence of an encrypted.xlsx file, loads it with the correct password via LoadOptions, reads cell A1 to confirm data integrity, and then saves the workbook as decrypted.xlsx without a password, while handling possible runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the encrypted workbook
            string encryptedFilePath = "encrypted.xlsx";

            // Verify the file exists to avoid FileNotFoundException
            if (!File.Exists(encryptedFilePath))
            {
                Console.WriteLine($"File not found: {encryptedFilePath}");
                return;
            }

            // Correct password for decryption
            string password = "myPassword";

            // LoadOptions with the password to open the encrypted workbook
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };

            // Load the workbook using the password
            Workbook workbook = new Workbook(encryptedFilePath, loadOptions);

            // Confirm data integrity by reading a known cell value (e.g., A1)
            Worksheet sheet = workbook.Worksheets[0];
            string cellValue = sheet.Cells["A1"].StringValue;
            Console.WriteLine("Decrypted cell A1 value: " + cellValue);

            // Save the workbook without a password to verify successful decryption
            string decryptedFilePath = "decrypted.xlsx";
            workbook.Save(decryptedFilePath, SaveFormat.Xlsx);
            Console.WriteLine($"Decrypted workbook saved to: {decryptedFilePath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
