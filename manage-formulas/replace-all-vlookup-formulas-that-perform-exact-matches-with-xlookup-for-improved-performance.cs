// Title: C# – Replace Exact‑Match VLOOKUP with XLOOKUP in Excel using Aspose.Cells
// Description: Load a workbook, detect VLOOKUP formulas that use FALSE for exact matching, convert them to XLOOKUP with INDEX helpers, recalculate, and save the updated file—all with Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# replace VLOOKUP | convert VLOOKUP to XLOOKUP .NET | Excel formula migration programmatically | regex VLOOKUP FALSE Aspose | bulk formula update Excel | XLOOKUP performance improvement | recalculate formulas Aspose.Cells | C# Excel automation
// Common Searches: replace VLOOKUP FALSE with XLOOKUP using Aspose.Cells C# | C# code to convert exact match VLOOKUP to XLOOKUP | bulk update Excel formulas from VLOOKUP to XLOOKUP | regex pattern for VLOOKUP exact match in .NET | Aspose.Cells example for formula replacement
// Developer Intent: Programmatically swap every exact‑match VLOOKUP formula for an equivalent XLOOKUP in an Excel workbook.
// Use Cases: Modernize legacy spreadsheets that rely on VLOOKUP(FALSE) for faster lookups. | Automate large‑scale formula conversion before generating reports or dashboards. | Ensure dependent calculations stay accurate by recalculating after formula changes.
// AI Prompts: Write C# code with Aspose.Cells that finds VLOOKUP formulas using FALSE and replaces them with XLOOKUP expressions. | Provide a regular expression and replacement logic to convert exact‑match VLOOKUP to XLOOKUP in an Excel file via Aspose.Cells. | Explain how to trigger formula recalculation after modifying cells with Aspose.Cells and save the workbook.

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Load a workbook, detect VLOOKUP formulas that use FALSE for exact matching, convert them to XLOOKUP with INDEX helpers, recalculate, and save the updated file—all with Aspose.Cells for .NET.
    public class Program
    {
        public static void Main()
        {
            try
            {
                ReplaceVlookupWithXlookup.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }

    public class ReplaceVlookupWithXlookup
    {
        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Regex to match VLOOKUP formulas with exact match (FALSE)
                Regex vlookupRegex = new Regex(
                    @"=VLOOKUP\(\s*([^,]+)\s*,\s*([^,]+)\s*,\s*([^,]+)\s*,\s*FALSE\s*\)",
                    RegexOptions.IgnoreCase);

                // Iterate through all worksheets and cells
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (Cell cell in sheet.Cells)
                    {
                        if (cell.IsFormula)
                        {
                            string formula = cell.Formula;

                            // Replace matching VLOOKUP with XLOOKUP
                            if (vlookupRegex.IsMatch(formula))
                            {
                                string newFormula = vlookupRegex.Replace(formula, match =>
                                {
                                    string lookupValue = match.Groups[1].Value.Trim();
                                    string tableArray = match.Groups[2].Value.Trim();
                                    string colIndex = match.Groups[3].Value.Trim();

                                    // Build XLOOKUP formula:
                                    // =XLOOKUP(lookup_value, INDEX(table_array,0,1), INDEX(table_array,0,col_index), ,0)
                                    return $"=XLOOKUP({lookupValue},INDEX({tableArray},0,1),INDEX({tableArray},0,{colIndex}),,0)";
                                });

                                cell.Formula = newFormula;
                            }
                        }
                    }
                }

                // Recalculate formulas after modifications
                workbook.CalculateFormula();

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}
