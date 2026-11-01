// Title: C# example: Remove unused cell styles from an Excel workbook using Aspose.Cells
// Common Searches: C# code to delete unused styles in an Excel file with Aspose.Cells | Aspose.Cells RemoveUnusedStyles method usage example | How to clean up Excel workbook styles programmatically in .NET | Load password‑protected Excel workbook and remove unused styles using Aspose.Cells | Save cleaned Excel file after removing unused styles with Aspose.Cells
// Tags: Workbook.RemoveUnusedStyles C# | load encrypted Excel workbook Aspose.Cells LoadOptions | save cleaned workbook Aspose.Cells | exception handling CellsException Aspose.Cells | optimize Excel file style footprint Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook (optionally using LoadOptions for a password), calls Workbook.RemoveUnusedStyles() to purge styles that are not referenced by any cell, and saves the cleaned file. It also demonstrates handling CellsException for password‑protected workbooks and generic exception handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook (add password to LoadOptions if required)
            var loadOptions = new LoadOptions(); // Set loadOptions.Password if needed
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Remove all styles that are not referenced by any cell
            workbook.RemoveUnusedStyles();

            // Save the cleaned workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (CellsException ex)
        {
            // Handle password‑protected workbook scenario
            if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("Error: The workbook is password protected. Provide a valid password in LoadOptions.");
            }
            else
            {
                Console.WriteLine($"Aspose.Cells error: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
