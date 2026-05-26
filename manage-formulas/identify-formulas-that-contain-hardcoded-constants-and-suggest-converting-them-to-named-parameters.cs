using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class IdentifyHardCodedConstants
    {
        public static void Main()
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
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists before attempting to load it.
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

            // Regex to locate numeric literals that are not part of cell references.
            Regex numberRegex = new Regex(@"(?<![A-Za-z$])(-?\d+(\.\d+)?)(?![A-Za-z])",
                                          RegexOptions.Compiled);

            // Iterate through each worksheet.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Iterate through all used cells.
                foreach (Cell cell in cells)
                {
                    if (!cell.IsFormula) continue;

                    string formula = cell.Formula;
                    MatchCollection matches = numberRegex.Matches(formula);

                    if (matches.Count == 0) continue;

                    Console.WriteLine($"Cell {cell.Name} has hard‑coded constants in formula: {formula}");

                    foreach (Match match in matches)
                    {
                        string constant = match.Value;
                        string paramName = $"Param_{constant.Replace(".", "_")}";

                        // Check if the named range already exists.
                        Name existingName = null;
                        try
                        {
                            existingName = workbook.Worksheets.Names[paramName];
                        }
                        catch { /* ignored – name does not exist */ }

                        if (existingName == null)
                        {
                            // Ensure a hidden sheet named "Parameters" exists.
                            Worksheet paramSheet = workbook.Worksheets["Parameters"];
                            if (paramSheet == null)
                            {
                                paramSheet = workbook.Worksheets.Add("Parameters");
                                paramSheet.IsVisible = false;
                            }

                            // Determine the next empty row.
                            int row = paramSheet.Cells.MaxDataRow + 1;
                            paramSheet.Cells[row, 0].PutValue(constant);

                            // Add a named range that points to the cell containing the constant.
                            int nameIndex = workbook.Worksheets.Names.Add(paramName);
                            Name name = workbook.Worksheets.Names[nameIndex];
                            name.RefersTo = $"=Parameters!{CellReference(row, 0)}";

                            Console.WriteLine($"  Created name '{paramName}' referring to {name.RefersTo}");
                        }
                        else
                        {
                            Console.WriteLine($"  Name '{paramName}' already exists.");
                        }

                        // Show a possible revised formula using the named parameter.
                        string revisedFormula = formula.Replace(constant, paramName);
                        Console.WriteLine($"  Suggested revised formula: {revisedFormula}");
                    }
                }
            }

            // Save the modified workbook.
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }

        // Convert zero‑based row/column indices to an A1‑style cell reference.
        private static string CellReference(int row, int column)
        {
            string colLetter = "";
            int col = column;
            do
            {
                colLetter = (char)('A' + (col % 26)) + colLetter;
                col = col / 26 - 1;
            } while (col >= 0);

            return $"{colLetter}{row + 1}";
        }
    }
}