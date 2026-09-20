// Title: C# utility to add missing leading '=' to formulas in an Excel workbook using Aspose.Cells
// AI Prompts: Write a C# method that loads an .xlsx file with Aspose.Cells, iterates every cell in all worksheets, and prepends '=' to any formula string that does not already start with it, then saves the workbook. | Create a console application that accepts input and output file paths, uses Aspose.Cells to detect cells containing formulas without a leading equal sign, corrects them in place, and writes the fixed workbook.
// Common Searches: aspnet c# scan excel workbook for formulas without leading equal sign using Aspose.Cells | how to automatically fix malformed formulas in .xlsx files with Aspose.Cells .NET | batch add leading '=' to Excel formulas that lack it via C# Aspose.Cells library | command line tool to correct missing leading equal sign in Excel formulas .NET
// Tags: add leading equal sign Aspose.Cells | detect malformed Excel formulas C# | batch correct formulas in .xlsx using Aspose.Cells | automate formula fixing with Aspose.Cells .NET | iterate worksheets to fix formula syntax

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsUtilities
{
    // The FormulaFixer class loads a workbook via Aspose.Cells, iterates through each worksheet's used range, checks each cell's Formula property, and if the formula is non‑empty and does not start with '=', prefixes it with '=', then saves the corrected workbook to the specified output path.
    public static class FormulaFixer
    {
        /// <param name="inputPath">Path to the source workbook.</param>
        /// <param name="outputPath">Path where the corrected workbook will be saved.</param>
        public static void FixMissingEqualSign(string inputPath, string outputPath)
        {
            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            try
            {
                // Load the workbook from the specified file.
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets in the workbook.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Get the used range of the worksheet to limit the iteration.
                    Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

                    // Iterate over each cell in the used range.
                    foreach (Cell cell in usedRange)
                    {
                        // Retrieve the formula string as stored in the cell.
                        string formula = cell.Formula;

                        // If the cell has a formula (non‑empty) and does not start with '=',
                        // it is considered a missing leading equal sign.
                        if (!string.IsNullOrEmpty(formula) && !formula.StartsWith("="))
                        {
                            // Prefix the formula with '=' and assign it back to the cell.
                            cell.Formula = "=" + formula;
                        }
                    }
                }

                // Save the modified workbook to the output path.
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                // Wrap and rethrow to provide context.
                throw new ApplicationException($"Error processing workbook '{inputPath}'.", ex);
            }
        }
    }

    class Program
    {
        /// <summary>
        /// Entry point for the console application.
        /// </summary>
        static void Main(string[] args)
        {
            // Example usage: provide input and output paths via command‑line arguments or defaults.
            string inputFile;
            string outputFile;

            if (args.Length >= 2)
            {
                inputFile = args[0];
                outputFile = args[1];
            }
            else
            {
                // Default paths for quick testing.
                inputFile = @"C:\Temp\Sample.xlsx";
                outputFile = @"C:\Temp\Sample_Fixed.xlsx";
            }

            try
            {
                FormulaFixer.FixMissingEqualSign(inputFile, outputFile);
                Console.WriteLine($"Workbook processed successfully. Output saved to: {outputFile}");
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.Error.WriteLine(fnfEx.Message);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
