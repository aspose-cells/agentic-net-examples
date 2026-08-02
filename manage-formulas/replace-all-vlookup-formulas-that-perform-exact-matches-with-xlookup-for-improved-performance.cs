// Title: Convert VLOOKUP(FALSE) to XLOOKUP in Excel using Aspose.Cells for .NET
// Description: A C# utility that loads an Excel file with Aspose.Cells, scans every worksheet for VLOOKUP formulas that use FALSE for exact matching, rewrites each to an equivalent XLOOKUP call (leveraging INDEX for lookup and return arrays), and saves the updated workbook.
// Keywords: Aspose.Cells | C# | .NET | VLOOKUP to XLOOKUP conversion | exact match lookup | Excel formula replacement | bulk formula update | XLOOKUP performance | programmatic Excel editing
// Common Searches: replace VLOOKUP FALSE with XLOOKUP Aspose.Cells C# | convert exact match VLOOKUP to XLOOKUP programmatically | bulk update Excel formulas using Aspose.Cells | regex VLOOKUP FALSE to XLOOKUP C# example | how to modernize VLOOKUP formulas in .NET
// Developer Intent: Automatically replace all VLOOKUP(...,FALSE) formulas in a workbook with equivalent XLOOKUP expressions using Aspose.Cells.
// Use Cases: Upgrade legacy spreadsheets to XLOOKUP for faster calculations. | Automate mass formula migration before distributing reports. | Prepare workbooks for newer Excel versions that favor XLOOKUP.
// AI Prompts: Generate C# code with Aspose.Cells that finds every VLOOKUP(...,FALSE) formula and substitutes it with XLOOKUP. | Provide a regex pattern and replacement logic to transform exact‑match VLOOKUP formulas to XLOOKUP in an Excel file. | Explain how to validate that XLOOKUP formulas produce the same results as the original VLOOKUP after conversion.

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // A C# utility that loads an Excel file with Aspose.Cells, scans every worksheet for VLOOKUP formulas that use FALSE for exact matching, rewrites each to an equivalent XLOOKUP call (leveraging INDEX for lookup and return arrays), and saves the updated workbook.
    public class ReplaceVlookupWithXlookup
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Regex to match VLOOKUP formulas with exact match (FALSE)
            Regex vlookupRegex = new Regex(
                @"VLOOKUP\(\s*([^,]+)\s*,\s*([^,]+)\s*,\s*([^,]+)\s*,\s*FALSE\s*\)",
                RegexOptions.IgnoreCase);

            // Iterate through all worksheets and cells
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                foreach (Cell cell in cells)
                {
                    if (!cell.IsFormula) continue;

                    string originalFormula = cell.Formula;

                    // Replace VLOOKUP with equivalent XLOOKUP
                    string updatedFormula = vlookupRegex.Replace(originalFormula, match =>
                    {
                        string lookupValue = match.Groups[1].Value.Trim();
                        string tableArray = match.Groups[2].Value.Trim();
                        string colIndex = match.Groups[3].Value.Trim();

                        // XLOOKUP(lookup_value, INDEX(table_array,0,1), INDEX(table_array,0,col_index), ,0)
                        return $"XLOOKUP({lookupValue},INDEX({tableArray},0,1),INDEX({tableArray},0,{colIndex}),,0)";
                    });

                    // Apply the new formula if it changed
                    if (!originalFormula.Equals(updatedFormula, StringComparison.Ordinal))
                    {
                        cell.Formula = updatedFormula;
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
