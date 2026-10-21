// Title: Bulk replace VLOOKUP formulas with INDEX‑MATCH in an Excel file using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells replace VLOOKUP formulas with INDEX MATCH in all worksheets | how to bulk convert VLOOKUP to INDEX-MATCH in an Excel workbook using Aspose.Cells | Aspose.Cells .NET code to update formulas from VLOOKUP to INDEX MATCH | optimize Excel lookup performance by swapping VLOOKUP for INDEX-MATCH with C#
// Tags: Aspose.Cells bulk formula replacement | convert VLOOKUP to INDEX-MATCH C# | Excel lookup performance optimization .NET | regex driven formula transformation Aspose | update Excel workbook formulas programmatically

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an input.xlsx workbook with Aspose.Cells, iterates through every worksheet and cell, detects VLOOKUP formulas using a regular expression, builds equivalent INDEX‑MATCH formulas, replaces the original formulas, and saves the modified workbook as output.xlsx, improving lookup performance.
class VlookupToIndexMatchConverter
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Regex to capture VLOOKUP arguments: lookup, table, column index, optional range_lookup
            Regex vlookupRegex = new Regex(
                @"VLOOKUP\s*\(\s*(?<lookup>[^,]+)\s*,\s*(?<table>[^,]+)\s*,\s*(?<col>\d+)\s*(?:,\s*(?<range>[^)]+))?\s*\)",
                RegexOptions.IgnoreCase);

            // Iterate through all worksheets and cells
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Cell cell in sheet.Cells)
                {
                    // Check if the cell contains a formula and includes VLOOKUP
                    if (cell.IsFormula && cell.Formula.IndexOf("VLOOKUP", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        string originalFormula = cell.Formula;

                        // Try to match the VLOOKUP pattern
                        Match match = vlookupRegex.Match(originalFormula);
                        if (match.Success)
                        {
                            // Extract and trim captured groups
                            string lookupValue = match.Groups["lookup"].Value.Trim();
                            string tableRange = match.Groups["table"].Value.Trim();
                            string colIndex = match.Groups["col"].Value.Trim();

                            // Build the INDEX‑MATCH formula
                            // =INDEX(tableRange, MATCH(lookupValue, INDEX(tableRange,0,1),0), colIndex)
                            string newFormula = $"=INDEX({tableRange}, MATCH({lookupValue}, INDEX({tableRange},0,1),0), {colIndex})";

                            // Replace the cell's formula
                            cell.Formula = newFormula;
                        }
                        // If the pattern does not match, leave the formula unchanged
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
