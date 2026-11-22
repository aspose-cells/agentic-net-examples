// Title: Decrypt an encrypted XLSX file and save it as an unencrypted workbook while preserving formulas and formatting with Aspose.Cells for .NET
// Common Searches: how to open a password‑protected XLSX in C# using Aspose.Cells LoadOptions | remove password from Excel file programmatically with Aspose.Cells .NET | save encrypted workbook as unencrypted while keeping formulas Aspose.Cells | C# code to decrypt an Excel file and retain cell formatting using Aspose.Cells | Aspose.Cells example for loading encrypted XLSX and exporting without password
// Tags: Aspose.Cells encrypted workbook loading | Aspose.Cells save workbook as unencrypted XLSX | Aspose.Cells preserve formulas during decryption | C# decrypt Excel file with Aspose.Cells | unprotected XLSX creation using Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample loads a password‑protected XLSX file using Aspose.Cells LoadOptions with the supplied password, then saves the workbook to a new XLSX file without specifying a password. All formulas, styles, and other content are retained, and the code includes basic error handling for missing files and invalid passwords.
class Program
{
    static void Main()
    {
        // Path to the encrypted XLSX file
        string encryptedFilePath = "encrypted.xlsx";

        // Password used to encrypt the source file
        string password = "yourPassword";

        // Verify that the encrypted file exists before attempting to load it
        if (!File.Exists(encryptedFilePath))
        {
            Console.WriteLine($"Error: The file \"{encryptedFilePath}\" was not found.");
            return;
        }

        try
        {
            // Load options with the password to open the encrypted workbook
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };

            // Load the encrypted workbook; formulas, formatting, and other content are preserved
            Workbook workbook = new Workbook(encryptedFilePath, loadOptions);

            // Path for the unencrypted output file
            string unencryptedFilePath = "unencrypted.xlsx";

            // Save the workbook without specifying a password (resulting file is unencrypted)
            workbook.Save(unencryptedFilePath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook has been saved unencrypted to \"{unencryptedFilePath}\".");
        }
        catch (Exception ex)
        {
            // Detect invalid password scenario based on exception message
            if (ex.Message != null && ex.Message.IndexOf("invalid password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("Error: The provided password is invalid for the encrypted workbook.");
            }
            else
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
