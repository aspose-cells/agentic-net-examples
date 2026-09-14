// Title: Bulk replace VLOOKUP formulas with INDEX‑MATCH in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that iterates through every worksheet, detects cells containing VLOOKUP, and rewrites each as an equivalent INDEX‑MATCH formula. | Extend the sample to handle VLOOKUP calls that omit the optional range_lookup argument, assuming an exact match by default. | Add error handling that skips cells whose VLOOKUP pattern cannot be parsed, logs the cell address, and continues processing the workbook.
// Common Searches: c# aspocells replace vlookup with index match across all sheets | how to convert VLOOKUP to INDEX-MATCH using Aspose.Cells .NET | bulk edit Excel formulas programmatically with Aspose.Cells C# | optimize Excel lookup performance by swapping VLOOKUP for INDEX-MATCH in C# | regex parse VLOOKUP formula Aspose.Cells example
// Tags: aspocells formula substitution vlookup | c# index-match conversion aspocells | bulk excel formula update .net | excel lookup speed improvement aspocells | regex parsing vlookup aspocells

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an Excel workbook, walks through every worksheet and cell, identifies formulas that contain VLOOKUP, extracts the arguments with a regular expression, builds an equivalent INDEX‑MATCH expression, replaces the original formula, and saves the updated workbook.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get all cells in the current worksheet
            Cells cells = sheet.Cells;

            // Loop through each cell that contains a formula
            foreach (Cell cell in cells)
            {
                if (cell.IsFormula && cell.Formula.IndexOf("VLOOKUP", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // Example VLOOKUP syntax:
                    // VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
                    // We'll replace it with an INDEX‑MATCH equivalent.

                    string formula = cell.Formula;

                    // Use a regular expression to capture the four VLOOKUP arguments.
                    // This simple pattern works for most straightforward cases.
                    Match vlookupMatch = Regex.Match(
                        formula,
                        @"VLOOKUP\(\s*(?<lookup>[^,]+)\s*,\s*(?<table>[^,]+)\s*,\s*(?<col>\d+)\s*,\s*(?<range>\w+)\s*\)",
                        RegexOptions.IgnoreCase);

                    if (vlookupMatch.Success)
                    {
                        string lookupValue = vlookupMatch.Groups["lookup"].Value.Trim();
                        string tableArray   = vlookupMatch.Groups["table"].Value.Trim();
                        string colIndexNum  = vlookupMatch.Groups["col"].Value.Trim();
                        string rangeLookup  = vlookupMatch.Groups["range"].Value.Trim();

                        // Determine match type: FALSE = exact match (0), TRUE or omitted = approximate (1)
                        int matchType = rangeLookup.Equals("FALSE", StringComparison.OrdinalIgnoreCase) ? 0 : 1;

                        // Build the INDEX‑MATCH formula.
                        // INDEX(table, MATCH(lookup, INDEX(table,0,1), match_type), col_index)
                        string indexMatchFormula = $"=INDEX({tableArray}, MATCH({lookupValue}, INDEX({tableArray},0,1), {matchType}), {colIndexNum})";

                        // Replace the original VLOOKUP formula with the new one.
                        cell.Formula = indexMatchFormula;
                    }
                }
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
