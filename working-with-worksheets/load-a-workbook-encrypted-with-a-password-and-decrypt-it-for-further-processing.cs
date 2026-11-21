// Title: Decrypt a password‑protected Excel workbook and save it unencrypted using Aspose.Cells for .NET
// Common Searches: how to load an encrypted .xlsx file with a password using Aspose.Cells in C# | remove password protection from an Excel workbook programmatically with Aspose.Cells .NET | Aspose.Cells example for opening a password‑protected workbook and saving it without a password | C# code to read a cell from a password‑protected Excel file using Aspose.Cells | Aspose.Cells LoadOptions Password property usage for decrypting Excel files
// Tags: Aspose.Cells LoadOptions password decryption | Workbook.Settings.Password clear operation | save decrypted Excel workbook Aspose.Cells | read cell A1 from encrypted workbook C# | verify file existence before loading Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookDecryptorApp
{
    // The sample verifies the source file, creates the output folder, loads the encrypted workbook with LoadOptions.Password, reads cell A1 from the first worksheet, clears the opening password via Workbook.Settings.Password, and saves the workbook as an unprotected .xlsx file.
    class WorkbookDecryptor
    {
        static void Main()
        {
            try
            {
                // Paths and password
                string inputPath = @"C:\Input\EncryptedWorkbook.xlsx";
                string outputPath = @"C:\Output\DecryptedWorkbook.xlsx";
                string password = "mySecretPassword";

                // Verify the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Load the encrypted workbook using the password
                LoadOptions loadOptions = new LoadOptions
                {
                    Password = password
                };
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Example: read the value of cell A1 from the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                string cellValue = sheet.Cells["A1"].StringValue;
                Console.WriteLine($"Value in A1: {cellValue}");

                // Remove the opening password (permanently decrypt)
                workbook.Settings.Password = string.Empty;

                // Save the decrypted workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved without password to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
