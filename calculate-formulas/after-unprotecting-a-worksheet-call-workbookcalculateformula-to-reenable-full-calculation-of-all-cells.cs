// Title: Unprotect an Excel worksheet and recalculate all formulas with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells remove worksheet protection and recalculate formulas | How to call Workbook.CalculateFormula after unprotecting a sheet in Aspose.Cells | Load password‑protected XLSX, unprotect sheet, then recalc all cells using Aspose.Cells | Example code for unprotecting a worksheet and triggering full formula calculation in C# | Aspose.Cells recalculate workbook after sheet unprotect without manual cell iteration
// Tags: unprotect worksheet Aspose.Cells C# | Workbook.CalculateFormula usage Aspose.Cells | load password protected xlsx Aspose.Cells | save recalculated workbook C# | worksheet protection removal Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads an XLSX workbook (optionally with a password), unprotects the first worksheet, runs Workbook.CalculateFormula to recompute every formula, and saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook. If the file is password‑protected, set the password here.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                // Set to the workbook password if required; leave empty for no password
                Password = ""
            };
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Access the first worksheet (or specify by name/index)
            Worksheet worksheet = workbook.Worksheets[0];

            // Unprotect the worksheet. Provide the sheet password if it is protected.
            worksheet.Unprotect(""); // use "" if the sheet has no password

            // Re‑calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Save the workbook after recalculation
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook processed and saved to: {outputPath}");
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
