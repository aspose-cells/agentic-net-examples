// Title: Load a password‑protected Excel workbook and save it without a password using Aspose.Cells for .NET
// Common Searches: C# load encrypted Excel file with Aspose.Cells and remove its password | how to save an Aspose.Cells workbook without password after opening | Aspose.Cells .NET clear workbook password after loading | remove workbook encryption using LoadOptions and Workbook.Settings in C# | unprotect Excel file programmatically with Aspose.Cells
// Tags: load password protected workbook Aspose.Cells | clear LoadOptions.Password C# | remove workbook encryption Aspose.Cells | save unprotected Excel file Aspose.Cells | Workbook.Settings.Password null C#

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to open an Excel workbook that is protected with a password using Aspose.Cells LoadOptions, clear the password from both LoadOptions and Workbook.Settings, and then save the workbook as an unprotected file, with error handling for CellsException and general exceptions.
class Program
{
    static void Main()
    {
        // Path to the password‑protected workbook
        string inputFile = "protected.xlsx";
        // Path where the unprotected workbook will be saved
        string outputFile = "unprotected.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file not found: {inputFile}");
                return;
            }

            // Set load options with the password required to open the file
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = "myPassword"
            };

            // Load the workbook using the provided password
            Workbook workbook = new Workbook(inputFile, loadOptions);

            // Clear the password from load options (optional cleanup)
            loadOptions.Password = null;

            // Remove encryption password from the workbook before saving
            workbook.Settings.Password = null;

            // Save the workbook without a password
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to {outputFile}");
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
