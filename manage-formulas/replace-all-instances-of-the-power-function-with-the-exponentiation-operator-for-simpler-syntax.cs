// Title: Replace POWER Function with ^ Operator in Excel Formulas Using Aspose.Cells for .NET
// Description: Loads a workbook, scans every worksheet and cell for formulas, uses a case‑insensitive regex to swap POWER(arg1,arg2) with (arg1)^(arg2), updates changed cells, recalculates all formulas, and saves the result.
// Keywords: Aspose.Cells replace POWER | Excel exponentiation operator C# | convert POWER to caret Aspose.Cells | regex formula transformation .NET | update workbook formulas programmatically
// Common Searches: replace POWER function with ^ in Aspose.Cells | C# change Excel POWER to caret operator | regex replace POWER(arg1,arg2) Aspose.Cells | modify all formulas in a workbook using Aspose.Cells
// Developer Intent: Automatically convert every POWER(arg1,arg2) formula in an Excel file to the ^ syntax with Aspose.Cells for .NET.
// Use Cases: Batch conversion of legacy POWER formulas before publishing a workbook | Standardizing formula syntax across large spreadsheets for compatibility | Pre‑processing worksheets to improve readability and reduce function calls
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells, replaces all POWER(arg1,arg2) formulas with (arg1)^(arg2) using a regex, recalculates, and saves the file. | Provide a regular expression pattern and replacement logic to convert POWER functions to the caret operator in an Aspose.Cells workbook. | Write a method that iterates through every cell in an Aspose.Cells workbook, detects POWER formulas, swaps them for exponentiation syntax, and updates the workbook.

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads a workbook, scans every worksheet and cell for formulas, uses a case‑insensitive regex to swap POWER(arg1,arg2) with (arg1)^(arg2), updates changed cells, recalculates all formulas, and saves the result.
    public class ReplacePowerFunctionDemo
    {
        public static void Run()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Ensure the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Regex to replace POWER(arg1,arg2) with (arg1)^(arg2)
                Regex powerRegex = new Regex(@"POWER\s*\(\s*([^,]+?)\s*,\s*([^\)]+?)\s*\)", RegexOptions.IgnoreCase);

                // Iterate through worksheets and cells
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (Cell cell in sheet.Cells)
                    {
                        if (cell.IsFormula)
                        {
                            string originalFormula = cell.Formula;
                            string updatedFormula = powerRegex.Replace(originalFormula, m =>
                            {
                                string left = m.Groups[1].Value.Trim();
                                string right = m.Groups[2].Value.Trim();
                                return $"({left})^({right})";
                            });

                            // Update the cell if the formula changed
                            if (!originalFormula.Equals(updatedFormula, StringComparison.Ordinal))
                            {
                                cell.Formula = updatedFormula;
                            }
                        }
                    }
                }

                // Recalculate all formulas
                workbook.CalculateFormula();

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Program entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            ReplacePowerFunctionDemo.Run();
        }
    }
}
