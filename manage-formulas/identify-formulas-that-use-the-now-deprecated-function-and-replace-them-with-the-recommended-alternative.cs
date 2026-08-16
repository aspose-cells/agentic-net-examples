// Title: C# – Replace Deprecated Excel Functions in a Workbook with Aspose.Cells
// Description: Loads a workbook, scans every worksheet and used cell, detects formulas that contain a deprecated function (e.g., OLD_FUNC), swaps it with the recommended replacement (NEW_FUNC) case‑insensitively, recalculates all formulas, and saves the updated file.
// Keywords: Aspose.Cells replace deprecated formula | C# bulk Excel function update | replace OLD_FUNC with NEW_FUNC | Aspose.Cells calculate formulas | .NET Excel formula migration | programmatic Excel function replacement | Aspose.Cells sample code | Excel workbook formula edit C#
// Common Searches: how to replace a deprecated Excel function using Aspose.Cells C# | bulk update formulas from OLD_FUNC to NEW_FUNC in .NET | Aspose.Cells recalculate workbook after formula changes | C# code to find and replace Excel functions in all cells | Aspose.Cells replace function name in formulas
// Developer Intent: Swap all occurrences of a removed Excel function with its modern equivalent and refresh calculations programmatically.
// Use Cases: Modernize legacy spreadsheets that still use obsolete functions before distribution. | Automate bulk conversion of multiple workbooks to the latest Excel standards. | Guarantee correct results after function replacement by invoking CalculateFormula.
// AI Prompts: Generate C# code with Aspose.Cells to find and replace a specific Excel function across an entire workbook. | Provide a step‑by‑step tutorial for replacing deprecated formulas and recalculating the workbook using Aspose.Cells. | Explain how to perform a case‑insensitive function name replacement in Excel formulas with Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads a workbook, scans every worksheet and used cell, detects formulas that contain a deprecated function (e.g., OLD_FUNC), swaps it with the recommended replacement (NEW_FUNC) case‑insensitively, recalculates all formulas, and saves the updated file.
    public class ReplaceDeprecatedFormulas
    {
        // Define the deprecated function name and its recommended replacement.
        private const string DeprecatedFunction = "OLD_FUNC";
        private const string ReplacementFunction = "NEW_FUNC";

        // Entry point for the console application.
        public static void Main(string[] args)
        {
            try
            {
                // Default file paths; can be overridden via command‑line arguments.
                string inputPath = args.Length > 0 ? args[0] : "input.xlsx";
                string outputPath = args.Length > 1 ? args[1] : "output.xlsx";

                Run(inputPath, outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }

        // Core processing method.
        public static void Run(string inputPath, string outputPath)
        {
            // Verify that the input workbook exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the existing workbook.
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets in the workbook.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;

                    // Iterate through all used cells.
                    foreach (Cell cell in cells)
                    {
                        // Process only cells that contain a formula.
                        if (cell.IsFormula && !string.IsNullOrEmpty(cell.Formula))
                        {
                            // Check if the formula uses the deprecated function.
                            if (cell.Formula.IndexOf(DeprecatedFunction, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                // Replace the deprecated function with the recommended alternative.
                                string updatedFormula = cell.Formula.Replace(DeprecatedFunction, ReplacementFunction, StringComparison.OrdinalIgnoreCase);

                                // Assign the updated formula back to the cell.
                                cell.Formula = updatedFormula;
                            }
                        }
                    }
                }

                // Recalculate all formulas after modifications.
                workbook.CalculateFormula();

                // Ensure the output directory exists.
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the updated workbook.
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
