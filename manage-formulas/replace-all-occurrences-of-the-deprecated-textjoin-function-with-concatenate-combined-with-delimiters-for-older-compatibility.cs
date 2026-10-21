// Title: Replace deprecated TEXTJOIN formulas with CONCATENATE in Excel workbooks using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells replace TEXTJOIN function with CONCATENATE in an existing workbook | How to convert TEXTJOIN formulas to CONCATENATE for older Excel versions using Aspose.Cells | Programmatically update Excel formulas: replace TEXTJOIN with CONCATENATE via Aspose.Cells .NET | Regex‑based formula transformation in Aspose.Cells to remove TEXTJOIN usage | Batch modify Excel workbook formulas from TEXTJOIN to CONCATENATE using C#
// Tags: Aspose.Cells formula conversion | C# Excel function replacement | batch modify Excel formulas | regex based formula editing | legacy Excel compatibility

using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, scans every worksheet and formula cell, uses a case‑insensitive regular expression to locate TEXTJOIN calls, extracts the delimiter and text arguments, builds an equivalent CONCATENATE formula by interleaving the delimiter, replaces the original formula, and saves the updated workbook to a new file for compatibility with older Excel versions.
class TextJoinReplacer
{
    static void Main(string[] args)
    {
        // Input and output file paths
        string inputPath = @"C:\Docs\SourceWorkbook.xlsx";
        string outputPath = @"C:\Docs\UpdatedWorkbook.xlsx";

        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Regex to capture TEXTJOIN arguments (case‑insensitive)
            Regex textJoinRegex = new Regex(@"TEXTJOIN\(([^)]*)\)", RegexOptions.IgnoreCase);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all cells in the sheet
                foreach (Cell cell in sheet.Cells)
                {
                    // Process only formula cells
                    if (cell.IsFormula)
                    {
                        string formula = cell.Formula; // Formula includes leading '='

                        // Look for TEXTJOIN in the formula
                        Match match = textJoinRegex.Match(formula);
                        if (match.Success)
                        {
                            // Arguments inside TEXTJOIN(...)
                            string argsInside = match.Groups[1].Value;

                            // Split on commas (simple split, assumes no nested commas)
                            string[] allArgs = argsInside
                                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                .Select(a => a.Trim())
                                .ToArray();

                            // TEXTJOIN requires at least three arguments: delimiter, ignore_empty, text1
                            if (allArgs.Length >= 3)
                            {
                                string delimiter = allArgs[0]; // First argument is the delimiter
                                // string ignoreEmpty = allArgs[1]; // Ignored for CONCATENATE conversion
                                string[] textArgs = allArgs.Skip(2).ToArray(); // Remaining texts to join

                                // Build CONCATENATE argument list interleaving the delimiter
                                List<string> newParts = new List<string>();
                                for (int i = 0; i < textArgs.Length; i++)
                                {
                                    newParts.Add(textArgs[i]); // Add the text argument
                                    if (i < textArgs.Length - 1)
                                        newParts.Add(delimiter); // Add delimiter between texts
                                }

                                // Construct the new formula string
                                string newFormula = "=CONCATENATE(" + string.Join(", ", newParts) + ")";

                                // Replace the original formula
                                cell.Formula = newFormula;
                            }
                        }
                    }
                }
            }

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred:");
            Console.WriteLine(ex.Message);
        }
    }
}
