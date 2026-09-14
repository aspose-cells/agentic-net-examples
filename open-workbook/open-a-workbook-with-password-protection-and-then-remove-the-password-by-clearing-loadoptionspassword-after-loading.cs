// Title: Load a password‑protected Excel workbook with Aspose.Cells for .NET and save it as an unencrypted file
// AI Prompts: Provide the workbook password via LoadOptions when constructing the Workbook object, then save the file without specifying a password. | After loading the file, call the unprotect function with the original password to release workbook protection before saving.
// Common Searches: Aspose.Cells .NET how to open encrypted Excel file and remove password | C# load password‑protected xlsx with LoadOptions and save as plain workbook | remove workbook structure protection using Aspose.Cells after loading encrypted file
// Tags: Aspose.Cells LoadOptions.Password for encrypted xlsx | Workbook.Unprotect method C# example | save workbook without password Aspose.Cells | open password protected Excel file Aspose.Cells .NET | clear workbook structure protection Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample checks for the existence of a password‑protected Excel file, loads it using LoadOptions.Password, optionally removes workbook structure protection, and then saves the workbook as an unencrypted file.
class Program
{
    static void Main()
    {
        // Path to the password‑protected workbook
        string inputPath = "protected.xlsx";
        string outputPath = "unprotected.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // LoadOptions with the password required to open the encrypted file
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = "myPassword"
            };

            // Load the workbook using the provided password
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // If the workbook itself is protected (structure), unprotect it
            // (use the same password if it was set for workbook protection)
            if (!string.IsNullOrEmpty(workbook.Settings.Password))
            {
                workbook.Unprotect("myPassword");
            }

            // Save the workbook without a password (no encryption)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved without password to: {outputPath}");
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
