// Title: How to open a password‑protected XLSX workbook with Aspose.Cells for .NET, remove its password, and save a decrypted copy
// AI Prompts: Load an encrypted .xlsx workbook using Aspose.Cells LoadOptions with a supplied password, clear the workbook's password property, and save the result as a new unprotected file. | Write C# code that opens a password‑protected Excel file via Aspose.Cells, removes the protection programmatically, and writes the decrypted workbook to disk.
// Common Searches: asp.net how to open encrypted Excel file with password using Aspose.Cells | c# remove password from xlsx using Aspose.Cells LoadOptions | decrypt password protected workbook programmatically Aspose.Cells | save unprotected copy of password protected Excel with Aspose.Cells .NET
// Tags: Aspose.Cells LoadOptions password decryption | clear workbook protection Aspose.Cells C# | save unencrypted XLSX Aspose.Cells | open encrypted Excel file .NET Aspose.Cells | programmatic Excel decryption .NET

using Aspose.Cells;
using System;
using System.IO;

// The example shows how to verify the encrypted XLSX file exists, configure LoadOptions with the correct password, load the workbook using Aspose.Cells, clear the workbook's password setting, and save the workbook as an unencrypted XLSX file, with handling for Aspose.Cells‑specific and general exceptions.
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

            // Configure load options with the password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };

            // Load the encrypted workbook
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Remove password protection
            workbook.Settings.Password = null;

            // Save the workbook without encryption
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Decrypted workbook saved to: {outputPath}");
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
