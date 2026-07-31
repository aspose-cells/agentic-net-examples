// Title: C# – Replace deprecated REPT with REPEAT in all Excel formulas using Aspose.Cells
// Description: Loads an existing workbook, iterates through every worksheet and cell, detects formulas that contain the REPT function, replaces each occurrence with the newer REPEAT function (case‑insensitive), and saves the updated workbook. Includes robust error handling for loading and saving.
// Keywords: Aspose.Cells | C# | REPT function | REPEAT function | Excel formula replace | bulk formula update | deprecated Excel function | worksheet iteration | programmatic Excel editing
// Common Searches: Aspose.Cells replace REPT with REPEAT | C# code to update Excel formulas programmatically | bulk replace function name in Excel workbook | case‑insensitive formula replace using Aspose.Cells | how to migrate REPT to REPEAT in Excel files
// Developer Intent: Replace every occurrence of the REPT function with REPEAT in all formulas of an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Modernize legacy spreadsheets that still use the deprecated REPT function. | Ensure compatibility with newer Excel versions that require REPEAT. | Automate large‑scale formula migrations across multiple worksheets before distribution.
// AI Prompts: Generate C# code with Aspose.Cells to replace a specific function name in all workbook formulas. | Show best‑practice error handling when loading, modifying, and saving Excel files with Aspose.Cells. | Explain how to perform a case‑insensitive function name replacement in Excel formulas using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads an existing workbook, iterates through every worksheet and cell, detects formulas that contain the REPT function, replaces each occurrence with the newer REPEAT function (case‑insensitive), and saves the updated workbook. Includes robust error handling for loading and saving.
    public class ReplaceReptWithRepeatDemo
    {
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
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Replace deprecated REPT function with REPEAT in all formulas
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                foreach (Cell cell in cells)
                {
                    if (cell.IsFormula)
                    {
                        string formula = cell.Formula;
                        if (!string.IsNullOrEmpty(formula) &&
                            formula.IndexOf("REPT", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            string updatedFormula = formula.Replace("REPT", "REPEAT", StringComparison.OrdinalIgnoreCase);
                            cell.Formula = updatedFormula;
                        }
                    }
                }
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
