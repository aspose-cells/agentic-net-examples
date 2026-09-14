// Title: Convert deprecated CHOOSE formulas to nested IF statements in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that scans all worksheets, finds formulas containing the CHOOSE function, and replaces each occurrence with an equivalent nested IF expression. | Implement a C# helper method that receives the argument list from a CHOOSE call and returns a properly formatted nested IF formula string for Aspose.Cells. | Add logging to the conversion script to output the original formula and the generated IF formula for every modified cell.
// Common Searches: aspnet replace CHOOSE function with IF in Excel using Aspose.Cells | how to transform CHOOSE formulas to nested IF in a .xlsx file programmatically | C# regex to locate and rewrite CHOOSE formulas in Aspose.Cells workbook | update deprecated Excel functions automatically with Aspose.Cells .NET
// Tags: CHOOSE to IF formula conversion Aspose.Cells | nested IF generation C# Aspose.Cells | regex formula replacement Excel .NET | deprecated Excel function migration Aspose.Cells | batch formula update workbook C#

using Aspose.Cells;
using System;
using System.IO;
using System.Text.RegularExpressions;

// Loads an input .xlsx file with Aspose.Cells, iterates through every worksheet and cell, detects formulas that contain the CHOOSE function, converts each CHOOSE call to an equivalent nested IF expression using a helper method, updates the cell formula, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Regex to locate CHOOSE functions in formulas
            var choosePattern = new Regex(@"CHOOSE\s*\(([^)]*)\)", RegexOptions.IgnoreCase);

            // Iterate through all worksheets and cells
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                var cells = sheet.Cells;
                foreach (Cell cell in cells)
                {
                    // Process only cells that contain a formula with CHOOSE
                    if (cell.IsFormula && cell.Formula.IndexOf("CHOOSE", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        string originalFormula = cell.Formula;

                        // Replace each CHOOSE occurrence with an equivalent nested IF formula
                        string newFormula = choosePattern.Replace(originalFormula, m => ConvertChooseToIf(m.Groups[1].Value));

                        // Assign the transformed formula back to the cell
                        cell.Formula = newFormula;
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

    /// <param name="args">The comma‑separated arguments inside CHOOSE()</param>
    /// <returns>Nested IF formula string</returns>
    static string ConvertChooseToIf(string args)
    {
        // Split arguments by commas (assumes no commas inside individual arguments)
        var parts = args.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length < 2)
            return "0"; // Fallback for malformed CHOOSE

        string indexExpression = parts[0].Trim(); // The index argument (e.g., A1)

        // Start with the last value as the default "else" part
        string nestedIf = parts[parts.Length - 1].Trim();

        // Build nested IFs from the second‑last value back to the first value
        for (int i = parts.Length - 2; i >= 1; i--)
        {
            string value = parts[i].Trim();
            // i corresponds to the position (1‑based) after the index argument
            nestedIf = $"IF({indexExpression}={i}, {value}, {nestedIf})";
        }

        return nestedIf;
    }
}
