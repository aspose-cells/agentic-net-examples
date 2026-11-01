// Title: Convert exact‑match VLOOKUP formulas to XLOOKUP in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# code to replace VLOOKUP FALSE lookup with XLOOKUP using Aspose.Cells | how to programmatically change Excel formulas from VLOOKUP to XLOOKUP in .NET | Aspose.Cells iterate all cells and update exact‑match VLOOKUP to XLOOKUP | bulk conversion of VLOOKUP formulas to XLOOKUP in a workbook with C# | replace VLOOKUP false range_lookup with XLOOKUP in an Excel file via Aspose API
// Tags: exact‑match VLOOKUP to XLOOKUP conversion Aspose.Cells | bulk Excel formula update C# | programmatic XLOOKUP generation with INDEX Aspose | worksheet cell iteration for formula replacement .NET | Excel workbook formula migration using Aspose.Cells | replace false range_lookup VLOOKUP Aspose API

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook, walks through every worksheet and cell, detects VLOOKUP formulas that use exact‑match (FALSE or 0), constructs an equivalent XLOOKUP formula using INDEX functions, replaces the original formula, and saves the updated file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        Workbook workbook = null;

        try
        {
            // Load the workbook; if the file is password‑protected, an exception will be thrown
            var loadOptions = new LoadOptions();
            workbook = new Workbook(inputPath, loadOptions);
        }
        catch (CellsException ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error while loading workbook: {ex.Message}");
            return;
        }

        try
        {
            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all cells in the worksheet
                foreach (Cell cell in sheet.Cells)
                {
                    // Process only cells that contain a VLOOKUP formula
                    if (cell.IsFormula && cell.Formula.Contains("VLOOKUP", StringComparison.OrdinalIgnoreCase))
                    {
                        string formula = cell.Formula;

                        // Extract the argument list inside VLOOKUP(...)
                        int startIdx = formula.IndexOf('(');
                        int endIdx = formula.LastIndexOf(')');
                        if (startIdx < 0 || endIdx < 0 || endIdx <= startIdx)
                            continue; // malformed formula, skip

                        string args = formula.Substring(startIdx + 1, endIdx - startIdx - 1);
                        // Simple split on commas (works for most cases without nested commas)
                        string[] parts = args.Split(',');

                        // VLOOKUP requires at least 4 arguments
                        if (parts.Length < 4)
                            continue; // not a standard VLOOKUP, skip

                        // Trim whitespace from each argument
                        for (int i = 0; i < parts.Length; i++)
                            parts[i] = parts[i].Trim();

                        string rangeLookup = parts[3];

                        // Replace only exact‑match lookups (FALSE or 0)
                        if (rangeLookup.Equals("FALSE", StringComparison.OrdinalIgnoreCase) ||
                            rangeLookup.Equals("0"))
                        {
                            string lookupValue = parts[0];
                            string tableArray = parts[1];
                            string colIndex = parts[2];

                            // Build XLOOKUP formula:
                            // =XLOOKUP(lookup_value, INDEX(table_array,0,1), INDEX(table_array,0,col_index), "#N/A", 0)
                            string newFormula = $"=XLOOKUP({lookupValue}, INDEX({tableArray},0,1), INDEX({tableArray},0,{colIndex}), \"#N/A\", 0)";

                            // Replace the VLOOKUP formula with the new XLOOKUP formula
                            cell.Formula = newFormula;
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (CellsException ex)
        {
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error during processing: {ex.Message}");
        }
    }
}
