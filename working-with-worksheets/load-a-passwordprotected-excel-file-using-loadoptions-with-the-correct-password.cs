// Title: Load a password‑protected Excel workbook with Aspose.Cells LoadOptions in C# and save it unencrypted
// AI Prompts: Generate C# code that opens a .xlsx file protected with a password using Aspose.Cells LoadOptions.Password, then saves the workbook without a password. | Demonstrate catching the specific CellsException thrown when the password supplied to LoadOptions does not match the workbook's encryption. | Provide an example that checks the source file exists, loads it with the correct password, and writes the unprotected workbook to a new location.
// Common Searches: aspnet load password protected xlsx using aspose.cells loadoptions c# | how to remove password from excel file programmatically with aspose.cells | c# catch cellsexception invalid password when opening encrypted workbook | aspose.cells load encrypted workbook and save without protection
// Tags: Aspose.Cells LoadOptions password | load encrypted xlsx C# | unprotect Excel workbook Aspose.Cells | handle CellsException invalid password | save workbook without encryption Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the presence of a password‑protected Excel file, configures LoadOptions with the correct password, loads the workbook using Aspose.Cells, and saves it to a new file without any password protection while handling both Aspose.Cells‑specific and general exceptions.
class Program
{
    static void Main()
    {
        // Path to the password‑protected workbook
        const string inputPath = "protected.xlsx";
        const string outputPath = "unprotected.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        // Password for the protected workbook
        string password = "myPassword";

        try
        {
            // Configure load options with the correct password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };

            // Load the protected workbook
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // (Optional) Perform any operations on the workbook here

            // Save the workbook without password protection
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
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
