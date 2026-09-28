// Title: How to open a password‑protected Excel workbook, remove its password, and save a decrypted copy using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells LoadOptions to open an encrypted .xlsx file with a given password, read a cell value, clear the workbook's opening password, and save the file without protection. | Generate C# code that loads a password‑protected workbook, removes its opening password, and writes the decrypted workbook to a new location using Aspose.Cells.
// Common Searches: asp.net load encrypted excel file with password using aspose.cells | remove password from excel workbook programmatically c# | example of LoadOptions.Password property in Aspose.Cells | decrypt an .xlsx file and save unprotected version c# | read cell A1 from password protected workbook using Aspose.Cells
// Tags: Aspose.Cells LoadOptions password decryption | C# clear opening password from Excel workbook | save unprotected .xlsx with Aspose.Cells | read cell value from encrypted workbook | error handling for loading encrypted Excel file Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the encrypted Excel file exists, loads it with the specified password via LoadOptions, reads cell A1 from the first worksheet, clears the workbook's opening password, and saves the decrypted workbook to a new file, with proper error handling for load and save operations.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the encrypted workbook
            string encryptedPath = "encrypted.xlsx";

            // Verify that the source file exists
            if (!File.Exists(encryptedPath))
            {
                Console.WriteLine($"Error: The file '{encryptedPath}' was not found.");
                return;
            }

            // Password used to protect the workbook
            string password = "yourPassword";

            // Load the workbook with the password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = password // provide password for decryption
            };

            Workbook workbook;
            try
            {
                workbook = new Workbook(encryptedPath, loadOptions);
            }
            catch (Exception loadEx)
            {
                Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                return;
            }

            // Example: read the value of cell A1 from the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            string cellValue = sheet.Cells["A1"].StringValue;
            Console.WriteLine($"Value of A1: {cellValue}");

            // Remove opening password before saving
            workbook.Settings.Password = string.Empty; // clears opening password

            // Save the decrypted workbook to a new file (or overwrite the original)
            string decryptedPath = "decrypted.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(decryptedPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            try
            {
                workbook.Save(decryptedPath);
                Console.WriteLine($"Workbook has been decrypted and saved to '{decryptedPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
