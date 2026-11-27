// Title: Decrypt a password‑protected XLSX workbook and save it as an unencrypted file using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# load encrypted Excel file with password | remove password from XLSX using Aspose.Cells .NET | how to decrypt a protected workbook programmatically in C# | save unprotected copy of password protected Excel with Aspose.Cells | load workbook with LoadOptions password and export without encryption
// Tags: Aspose.Cells LoadOptions with password | decrypt encrypted XLSX workbook | clear workbook password setting | save unprotected Excel file C# | open password protected workbook Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example checks for the encrypted file, loads it using LoadOptions that include the correct password, clears the workbook's password property to remove encryption, and saves a new XLSX file without protection, handling Aspose.Cells‑specific and general exceptions.
class Program
{
    static void Main()
    {
        // Path to the encrypted workbook
        string inputPath = "encrypted.xlsx";

        // Password used to encrypt the workbook
        string password = "myPassword";

        // Path for the decrypted output workbook
        string outputPath = "decrypted.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the encrypted workbook with the correct password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Remove any password before saving to produce an unencrypted file
            workbook.Settings.Password = string.Empty; // clears encryption

            // Save the workbook as a new XLSX file without password protection
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Decrypted workbook saved to: {outputPath}");
        }
        catch (CellsException ex)
        {
            // Handles Aspose.Cells specific errors (e.g., invalid password)
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handles any other unexpected errors
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
