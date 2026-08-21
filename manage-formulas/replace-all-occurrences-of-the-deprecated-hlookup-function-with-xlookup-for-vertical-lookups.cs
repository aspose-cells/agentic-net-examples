// Title: C# – Replace HLOOKUP with XLOOKUP in Excel using Aspose.Cells
// Description: Loads a workbook, scans every worksheet for HLOOKUP formulas, swaps them to XLOOKUP (case‑insensitive), recalculates all formulas, and saves the updated file.
// Keywords: Aspose.Cells replace HLOOKUP | convert HLOOKUP to XLOOKUP .NET | bulk formula update C# | Excel XLOOKUP migration | programmatic Excel formula replacement | Aspose.Cells formula recalculation
// Common Searches: how to change HLOOKUP to XLOOKUP with Aspose.Cells | C# code to replace deprecated Excel functions | bulk update of Excel formulas using Aspose.Cells | automate HLOOKUP to XLOOKUP conversion .NET | Aspose.Cells replace formula text
// Developer Intent: Programmatically replace every HLOOKUP formula in an Excel workbook with an XLOOKUP formula using Aspose.Cells for .NET.
// Use Cases: Modernize legacy spreadsheets before distribution. | Automate mass conversion of workbooks during a version upgrade. | Ensure compatibility with newer Excel versions by updating deprecated functions. | Recalculate sheets after formula changes to validate results.
// AI Prompts: Write C# code with Aspose.Cells that opens a workbook, finds all cells containing HLOOKUP, replaces the function name with XLOOKUP, recalculates the workbook, and saves it. | Provide a robust C# routine that parses HLOOKUP arguments and builds equivalent XLOOKUP syntax, handling optional parameters and error checking.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads a workbook, scans every worksheet for HLOOKUP formulas, swaps them to XLOOKUP (case‑insensitive), recalculates all formulas, and saves the updated file.
    public class ReplaceHlookupWithXlookup
    {
        // Entry point for the console application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Define input and output file paths (adjust as needed)
            string inputPath = "InputWorkbook.xlsx";
            string outputPath = "OutputWorkbook.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;

                    // Loop through all used cells
                    foreach (Cell cell in cells)
                    {
                        // Process only cells that contain a formula
                        if (cell.IsFormula)
                        {
                            string formula = cell.Formula;

                            // Check if the formula uses the deprecated HLOOKUP function
                            if (!string.IsNullOrEmpty(formula) &&
                                formula.IndexOf("HLOOKUP", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                // Replace HLOOKUP with XLOOKUP (simple textual replacement)
                                string newFormula = formula.Replace("HLOOKUP", "XLOOKUP", StringComparison.OrdinalIgnoreCase);
                                cell.Formula = newFormula;
                            }
                        }
                    }
                }

                // Recalculate all formulas after the replacements
                workbook.CalculateFormula();

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}
