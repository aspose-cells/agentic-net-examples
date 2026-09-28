// Title: Decrypt an encrypted XLSX file and save it as an unprotected workbook while preserving formulas and formatting with Aspose.Cells for .NET
// AI Prompts: Load a password‑protected XLSX using Aspose.Cells LoadOptions, then save it without a password to keep all formulas and styles. | Write C# code that opens an encrypted Excel workbook with a given password and outputs a plain XLSX file that retains worksheet content. | Show how to remove workbook encryption in Aspose.Cells while ensuring formulas, charts, and formatting remain intact.
// Common Searches: asp.net how to open encrypted xlsx with password using aspose.cells | c# remove password from excel file without losing formulas | convert protected xlsx to unprotected using Aspose.Cells LoadOptions | preserve cell formatting when decrypting an Excel workbook in .NET | Aspose.Cells decrypt workbook and save as plain xlsx example
// Tags: Aspose.Cells load encrypted XLSX | Aspose.Cells save workbook without password | remove Excel file encryption C# | preserve formulas when decrypting Excel | LoadOptions password protected workbook

using Aspose.Cells;
using System;
using System.IO;

// // Opens an encrypted XLSX using LoadOptions with the supplied password, then saves it as a new XLSX without a password, preserving all formulas, styles, and other worksheet content.
class Program
{
    static void Main()
    {
        // Path to the encrypted XLSX file
        string encryptedFilePath = "encrypted.xlsx";

        // Verify that the encrypted file exists
        if (!File.Exists(encryptedFilePath))
        {
            Console.WriteLine($"Error: The file \"{encryptedFilePath}\" was not found.");
            return;
        }

        // Password used to encrypt the file
        string password = "yourPassword";

        try
        {
            // Load the encrypted workbook using LoadOptions with the password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };
            Workbook workbook = new Workbook(encryptedFilePath, loadOptions);

            // Path for the unencrypted output file
            string unencryptedFilePath = "unencrypted.xlsx";

            // Save the workbook without a password (unencrypted)
            // Using SaveFormat directly avoids version‑specific SaveOptions classes
            workbook.Save(unencryptedFilePath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved successfully to \"{unencryptedFilePath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
