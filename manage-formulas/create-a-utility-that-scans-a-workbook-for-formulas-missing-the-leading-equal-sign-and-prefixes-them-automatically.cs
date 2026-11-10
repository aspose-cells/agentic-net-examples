// Title: C# utility that scans an Excel workbook and automatically prefixes missing leading '=' to formula text using Aspose.Cells
// Common Searches: C# code to add a leading equal sign to Excel formula strings using Aspose.Cells | how to batch fix missing '=' in Excel formulas with Aspose.Cells .NET | scan all worksheets and convert text that looks like a formula into a real formula in C# | regex pattern to detect Excel function syntax and prepend '=' via Aspose.Cells | automate correction of formula cells lacking '=' in .xlsx files with Aspose.Cells
// Tags: batch prepend equal sign to Excel formulas Aspose.Cells | detect formula-like text with regex in .xlsx using C# | convert string cells to formulas Aspose.Cells .NET | iterate worksheets and cells to fix missing '=' Aspose.Cells | automated Excel formula correction utility C#

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsUtilities
{
    // Loads a workbook, iterates through every worksheet and used cell, identifies string cells that match a function‑style pattern but lack a leading '=', prefixes the '=', assigns the value as a formula, and saves the corrected workbook to the specified output path.
    public static class FormulaFixer
    {
        /// <param name="inputPath">Path to the source workbook.</param>
        /// <param name="outputPath">Path where the corrected workbook will be saved.</param>
        public static void FixMissingEqualSign(string inputPath, string outputPath)
        {
            // Load the workbook (provided load rule)
            Workbook workbook = new Workbook(inputPath);

            // Regular expression to identify text that looks like a formula (e.g., SUM(A1:B2))
            // It matches a function name followed by parentheses.
            Regex formulaPattern = new Regex(@"^[A-Za-z_][A-Za-z0-9_]*\s*\(.*\)$", RegexOptions.Compiled);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all used cells in the worksheet
                foreach (Cell cell in sheet.Cells)
                {
                    // Process only string cells
                    if (cell.Type == CellValueType.IsString)
                    {
                        string text = cell.StringValue;

                        // Skip if the text already starts with '='
                        if (text.StartsWith("="))
                            continue;

                        // If the text matches the formula pattern, treat it as a missing‑equal formula
                        if (formulaPattern.IsMatch(text))
                        {
                            // Prefix the missing equal sign and assign as a formula
                            cell.Formula = "=" + text;
                        }
                    }
                }
            }

            // Save the corrected workbook (provided save rule)
            workbook.Save(outputPath);
        }

        // Entry point for the console application
        public static void Main(string[] args)
        {
            try
            {
                string inputFile;
                string outputFile;

                if (args.Length >= 2)
                {
                    inputFile = args[0];
                    outputFile = args[1];
                }
                else
                {
                    // Default sample paths – adjust as needed
                    inputFile = @"C:\Temp\Sample.xlsx";
                    outputFile = @"C:\Temp\Sample_Fixed.xlsx";
                }

                // Verify that the input file exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Error: Input file not found – {inputFile}");
                    return;
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputFile);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Perform the formula fix
                FixMissingEqualSign(inputFile, outputFile);
                Console.WriteLine("Formula correction completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
