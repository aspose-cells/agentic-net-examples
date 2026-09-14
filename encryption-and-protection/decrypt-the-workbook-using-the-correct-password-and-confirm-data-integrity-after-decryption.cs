// Title: Decrypt a password‑protected Excel (.xlsx) workbook with Aspose.Cells for .NET and verify cell A1 value
// AI Prompts: Use Aspose.Cells LoadOptions.Password to open an encrypted .xlsx file, read a specific cell, and save the workbook without a password. | Check the existence of the source file, load it with the provided password, output the value of cell A1 to confirm decryption, and write the unprotected workbook to a new file. | Implement try‑catch blocks for CellsException and generic Exception to handle decryption errors and report meaningful messages.
// Common Searches: how to open a password protected Excel file using Aspose.Cells C# | Aspose.Cells decrypt workbook and remove password programmatically | verify cell content after loading encrypted .xlsx with Aspose.Cells | C# load encrypted Excel with password and save as unprotected file | Aspose.Cells LoadOptions Password property example
// Tags: Aspose.Cells password decryption via LoadOptions | C# remove password from Excel workbook | read cell after workbook decryption Aspose.Cells | save unprotected .xlsx with Aspose.Cells | handle CellsException during decryption

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the encrypted Excel file, configures LoadOptions with the correct password, loads the workbook, reads cell A1 to confirm successful decryption, and saves the workbook as an unprotected .xlsx file while handling Aspose.Cells‑specific and generic exceptions.
class Program
{
    static void Main()
    {
        const string inputFile = "EncryptedWorkbook.xlsx";
        const string outputFile = "DecryptedWorkbook.xlsx";
        const string password = "MySecretPassword"; // replace with actual password

        // Verify that the encrypted workbook exists
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: File \"{inputFile}\" not found.");
            return;
        }

        try
        {
            // Set load options with the correct password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };

            // Load the encrypted workbook
            Workbook workbook = new Workbook(inputFile, loadOptions);

            // Read a known cell to confirm successful decryption
            string decryptedValue = workbook.Worksheets[0].Cells["A1"].StringValue;
            Console.WriteLine("Decrypted cell A1 value: " + decryptedValue);

            // Save the decrypted workbook to a new file
            workbook.Save(outputFile);
            Console.WriteLine($"Decrypted workbook saved as \"{outputFile}\".");
        }
        catch (CellsException ex)
        {
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
