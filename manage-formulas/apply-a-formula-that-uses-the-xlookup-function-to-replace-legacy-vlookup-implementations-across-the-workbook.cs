// Title: Replace VLOOKUP formulas with XLOOKUP in every worksheet of an Excel file using Aspose.Cells for .NET
// AI Prompts: Load an .xlsx workbook with Aspose.Cells, loop through all worksheets, find cells that contain VLOOKUP, and rewrite each formula to an XLOOKUP expression in C#. | Use a regular expression to extract the lookup value, table array, and column index from a VLOOKUP formula, build the matching XLOOKUP syntax, and assign it back to the cell via the Aspose.Cells API. | After converting all lookup formulas, save the updated workbook to a new file while preserving the original formatting and data.
// Common Searches: how to convert VLOOKUP to XLOOKUP programmatically with Aspose.Cells C# | replace legacy lookup formulas in all sheets using Aspose.Cells .NET | regex extract arguments from Excel formula in C# Aspose.Cells example | bulk update Excel formulas across workbook using Aspose.Cells library
// Tags: Aspose.Cells replace VLOOKUP with XLOOKUP | C# bulk formula transformation Excel workbook | regex parse Excel formula arguments .NET | update cell formulas across worksheets Aspose.Cells | save modified workbook Aspose.Cells .xlsx

using System;
using System.IO;
using Aspose.Cells;
using System.Text.RegularExpressions;

namespace AsposeCellsExample
{
    // The example loads an input .xlsx file, iterates through every worksheet and cell, detects VLOOKUP formulas, extracts their arguments with a regular expression, constructs equivalent XLOOKUP formulas, replaces the original formulas, and saves the workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "InputWorkbook.xlsx";
            string outputPath = "OutputWorkbook.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;
                    foreach (Cell cell in cells)
                    {
                        // Process only cells that contain formulas
                        if (cell.IsFormula)
                        {
                            string formula = cell.Formula;

                            // Look for VLOOKUP usage
                            if (formula.IndexOf("VLOOKUP", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                Match match = Regex.Match(
                                    formula,
                                    @"VLOOKUP\s*\(\s*(?<args>.+?)\s*\)",
                                    RegexOptions.IgnoreCase);

                                if (match.Success)
                                {
                                    // Split arguments (simple split works for typical VLOOKUP)
                                    string[] formulaArgs = match.Groups["args"].Value.Split(',');

                                    if (formulaArgs.Length >= 3)
                                    {
                                        for (int i = 0; i < formulaArgs.Length; i++)
                                            formulaArgs[i] = formulaArgs[i].Trim();

                                        string lookupValue = formulaArgs[0];
                                        string tableArray = formulaArgs[1];
                                        string colIndexNum = formulaArgs[2];

                                        // Build equivalent XLOOKUP formula
                                        string xlookupFormula = $"=XLOOKUP({lookupValue},INDEX({tableArray},0,1),INDEX({tableArray},0,{colIndexNum}),\"\",0,1)";

                                        // Replace the old VLOOKUP formula
                                        cell.Formula = xlookupFormula;
                                    }
                                }
                            }
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
