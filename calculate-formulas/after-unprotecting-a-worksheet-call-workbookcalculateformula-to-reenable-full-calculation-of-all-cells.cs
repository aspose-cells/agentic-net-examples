// Title: Unprotect an Excel worksheet and recalculate all formulas with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells unprotect worksheet then recalculate formulas | How to call Workbook.CalculateFormula after sheet.Unprotect in .NET | Load password‑protected XLSX with Aspose.Cells and refresh formulas | Aspose.Cells example to remove sheet protection and recalc all cells | Re‑enable full calculation after unprotecting a sheet using Aspose.Cells C#
// Tags: worksheet unprotect Aspose.Cells C# | Workbook.CalculateFormula usage | load password protected XLSX Aspose.Cells | recalculate all formulas after unprotect | Aspose.Cells exception handling example

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // This C# example loads an optional password‑protected XLSX file with Aspose.Cells, unprotects the first worksheet, triggers full formula recalculation via Workbook.CalculateFormula, and saves the result, demonstrating proper error handling for both Aspose.Cells‑specific and general exceptions.
    class Program
    {
        static void Main(string[] args)
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
                // If the workbook is password‑protected, set the password here.
                // Leave empty ("") if no password is required.
                var loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    Password = "" // TODO: replace with actual password if needed
                };

                // Load the workbook with the specified options
                var workbook = new Workbook(inputPath, loadOptions);

                // Access the first worksheet
                var sheet = workbook.Worksheets[0];

                // If the worksheet is protected, provide the password; otherwise this call is safe.
                sheet.Unprotect(); // TODO: sheet.Unprotect("sheetPassword") if needed

                // Re‑enable full calculation of all cells
                workbook.CalculateFormula();

                // Save the processed workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (CellsException ex)
            {
                // Handles Aspose.Cells specific errors (e.g., wrong password)
                Console.WriteLine($"Aspose.Cells error: {ex.Message}");
            }
            catch (Exception ex)
            {
                // Handles any other unexpected errors
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
