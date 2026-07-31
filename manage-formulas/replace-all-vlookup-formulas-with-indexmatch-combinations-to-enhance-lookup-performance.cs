// Title: Bulk replace VLOOKUP with INDEX‑MATCH in Excel using Aspose.Cells for .NET (C#)
// Description: A C# utility that loads an Excel workbook with Aspose.Cells, scans every worksheet for VLOOKUP formulas using a tolerant regex, converts each to an equivalent INDEX‑MATCH expression, updates the cell, recalculates all formulas, and saves the optimized file.
// Keywords: Aspose.Cells | C# | .NET | VLOOKUP replacement | INDEX MATCH conversion | bulk formula update | Excel performance | regex formula detection | workbook automation | Excel lookup optimization
// Common Searches: replace VLOOKUP with INDEX MATCH Aspose.Cells C# | bulk convert Excel formulas using Aspose.Cells | regex to find VLOOKUP in .NET workbook | optimize Excel lookup speed with Aspose.Cells | C# code to change all VLOOKUP formulas
// Developer Intent: Automatically transform every VLOOKUP formula in a workbook into an INDEX‑MATCH equivalent to improve calculation speed.
// Use Cases: Migrate legacy spreadsheets to faster lookup logic before distribution. | Accelerate large financial models by swapping VLOOKUP for INDEX‑MATCH across all sheets. | Provide a server‑side service that receives user Excel files, rewrites lookup formulas, and returns an optimized version.
// AI Prompts: Generate C# code with Aspose.Cells that detects VLOOKUP formulas via regex and replaces them with INDEX‑MATCH while keeping absolute references intact. | Create unit tests that verify the VLOOKUP‑to‑INDEX‑MATCH conversion works for exact match, approximate match, and named‑range scenarios. | Explain how to extend the regex to capture VLOOKUP calls that include sheet‑qualified ranges or named ranges.

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsVlookupReplacement
{
    // A C# utility that loads an Excel workbook with Aspose.Cells, scans every worksheet for VLOOKUP formulas using a tolerant regex, converts each to an equivalent INDEX‑MATCH expression, updates the cell, recalculates all formulas, and saves the optimized file.
    class Program
    {
        static void Main()
        {
            // Load the workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Regular expression to capture VLOOKUP arguments:
            // =VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
            // This pattern is tolerant to spaces and optional fourth argument.
            Regex vlookupRegex = new Regex(
                @"=VLOOKUP\s*\(\s*(?<lookup>[^,]+)\s*,\s*(?<table>[^,]+)\s*,\s*(?<col>\d+)\s*(,\s*(?<range>TRUE|FALSE))?\s*\)",
                RegexOptions.IgnoreCase);

            // Iterate through all worksheets and cells
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                foreach (Cell cell in cells)
                {
                    // Process only cells that contain a formula
                    if (cell.IsFormula)
                    {
                        string formula = cell.Formula;

                        // Check if the formula contains VLOOKUP
                        Match match = vlookupRegex.Match(formula);
                        if (match.Success)
                        {
                            // Extract parts of the VLOOKUP formula
                            string lookupValue = match.Groups["lookup"].Value.Trim();
                            string tableArray = match.Groups["table"].Value.Trim();
                            string colIndex = match.Groups["col"].Value.Trim();
                            // range_lookup (TRUE/FALSE) is ignored because INDEX-MATCH performs exact match by default
                            
                            // Build the equivalent INDEX-MATCH formula
                            // =INDEX(table_array, MATCH(lookup_value, INDEX(table_array,0,1), 0), col_index_num)
                            string indexMatchFormula = $"=INDEX({tableArray}, MATCH({lookupValue}, INDEX({tableArray},0,1), 0), {colIndex})";

                            // Replace the original VLOOKUP formula with the new one
                            cell.Formula = indexMatchFormula;
                        }
                    }
                }
            }

            // Recalculate all formulas after replacement
            workbook.CalculateFormula();

            // Save the modified workbook (replace with desired output path)
            workbook.Save("output.xlsx");
        }
    }
}
