// Title: Convert VLOOKUP to INDEX‑MATCH in Excel with Aspose.Cells for .NET
// Description: C# program that loads an Excel workbook using Aspose.Cells, scans every worksheet for VLOOKUP formulas, parses their arguments, builds equivalent INDEX‑MATCH expressions (preserving exact or approximate match), replaces the original formulas, forces a full recalculation, and saves the updated file. Improves lookup speed and modernizes legacy spreadsheets.
// Keywords: Aspose.Cells VLOOKUP replace | INDEX MATCH conversion .NET | C# Excel formula update | bulk formula replacement | Excel performance optimization | programmatic formula rewrite | Aspose.Cells workbook manipulation
// Common Searches: replace VLOOKUP with INDEX MATCH using Aspose.Cells | C# code to convert Excel VLOOKUP formulas | how to update formulas in Aspose.Cells workbook | recalculate workbook after formula changes Aspose.Cells | bulk VLOOKUP to INDEX MATCH conversion .NET
// Developer Intent: Automatically replace all VLOOKUP formulas in an Excel workbook with faster INDEX‑MATCH equivalents via Aspose.Cells.
// Use Cases: Modernizing legacy spreadsheets before distribution to end users. | Optimizing uploaded Excel files in a web service for faster calculation. | Ensuring consistent, high‑performance lookup logic across multiple worksheets.
// AI Prompts: Write C# code with Aspose.Cells that scans a workbook, detects VLOOKUP formulas, parses arguments (including quoted commas), and substitutes them with correct INDEX‑MATCH formulas. | Provide a robust VLOOKUP argument parser that handles nested functions and quoted strings, then generates the matching INDEX‑MATCH expression. | Explain how to trigger a full workbook recalculation and safely save the file after programmatically modifying formulas with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace VlookupToIndexMatch
{
    // C# program that loads an Excel workbook using Aspose.Cells, scans every worksheet for VLOOKUP formulas, parses their arguments, builds equivalent INDEX‑MATCH expressions (preserving exact or approximate match), replaces the original formulas, forces a full recalculation, and saves the updated file. Improves lookup speed and modernizes legacy spreadsheets.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all cells that contain formulas
                    foreach (Cell cell in sheet.Cells)
                    {
                        if (cell.IsFormula)
                        {
                            string formula = cell.Formula;

                            // Check if the formula is a VLOOKUP (case‑insensitive)
                            if (formula.StartsWith("=VLOOKUP", StringComparison.OrdinalIgnoreCase))
                            {
                                // Extract the argument list inside the parentheses
                                int openParen = formula.IndexOf('(');
                                int closeParen = formula.LastIndexOf(')');
                                if (openParen < 0 || closeParen < 0 || closeParen <= openParen)
                                    continue; // malformed formula, skip

                                string argsInside = formula.Substring(openParen + 1, closeParen - openParen - 1);
                                // Simple split by commas (does not handle commas inside quoted strings)
                                string[] parts = argsInside.Split(',');

                                if (parts.Length < 3)
                                    continue; // not enough arguments, skip

                                // Trim whitespace from each part
                                string lookupValue = parts[0].Trim();
                                string tableArray = parts[1].Trim();
                                string colIndex = parts[2].Trim();
                                string rangeLookup = parts.Length > 3 ? parts[3].Trim() : "FALSE";

                                // Determine match_type for MATCH: 0 = exact (FALSE), 1 = approximate (TRUE)
                                string matchType = rangeLookup.Equals("TRUE", StringComparison.OrdinalIgnoreCase) ? "1" : "0";

                                // Build the INDEX‑MATCH formula
                                // =INDEX(tableArray, MATCH(lookupValue, INDEX(tableArray,0,1), matchType), colIndex)
                                string newFormula = $"=INDEX({tableArray}, MATCH({lookupValue}, INDEX({tableArray},0,1), {matchType}), {colIndex})";

                                // Replace the VLOOKUP formula with the new INDEX‑MATCH formula
                                cell.Formula = newFormula;
                            }
                        }
                    }
                }

                // Recalculate all formulas after replacement
                workbook.CalculateFormula();

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
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
}
