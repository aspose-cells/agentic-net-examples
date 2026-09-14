// Title: Decrypt a password‑protected XLSX file with Aspose.Cells LightCells API in C#
// AI Prompts: Generate C# code that loads a password‑protected XLSX workbook using Aspose.Cells LightCells API by setting the Password property in LoadOptions, then saves it as an unencrypted file. | Write a C# snippet that verifies the existence of a protected Excel file, opens it with LoadOptions containing the decryption password, and writes the decrypted workbook to a new location. | Create robust C# error handling for opening an encrypted workbook with Aspose.Cells LightCells API and exporting it without password protection.
// Common Searches: c# aspose.cells load encrypted xlsx with password | how to remove password from excel file using aspose cells lightcells | load protected workbook using loadoptions password property aspose.cells | decrypt xlsx file programmatically aspose.cells c# example | aspose.cells lightcells api open password protected workbook
// Tags: Aspose.Cells LightCells open encrypted workbook | C# LoadOptions set password | Excel workbook decryption Aspose.Cells | unprotected workbook save Aspose.Cells | handling protected Excel files LightCells

using System;
using System.IO;
using Aspose.Cells;

namespace DecryptWorkbook
{
    // // This program checks for a password‑protected XLSX file, loads it with the specified decryption password via LoadOptions (LightCells API), and saves the workbook as an unencrypted XLSX file.
    class Program
    {
        static void Main()
        {
            // Path to the password‑protected XLSX file
            string inputPath = "protected.xlsx";

            // Path where the decrypted workbook will be saved
            string outputPath = "decrypted.xlsx";

            // Decryption password
            string password = "myPassword";

            try
            {
                // Ensure the input file exists
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

                // Load the protected workbook using the password
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Save the workbook (now decrypted) to a new file
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Decrypted workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
