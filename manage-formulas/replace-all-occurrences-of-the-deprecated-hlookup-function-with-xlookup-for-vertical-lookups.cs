// Title: Programmatically replace HLOOKUP formulas with XLOOKUP in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells replace HLOOKUP formulas with XLOOKUP in existing workbook | how to convert HLOOKUP to XLOOKUP using Aspose.Cells .NET API | bulk update Excel formulas from HLOOKUP to XLOOKUP with Aspose.Cells | replace vertical lookup function in Excel file programmatically using Aspose.Cells | use regex to find and change HLOOKUP to XLOOKUP in C# Aspose.Cells
// Tags: replace HLOOKUP with XLOOKUP Aspose.Cells | bulk formula conversion .NET | Excel formula migration C# | regex formula parsing Aspose.Cells | vertical lookup update Excel .NET

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads a workbook, scans every worksheet and cell for HLOOKUP formulas, extracts the lookup value, table array, and row index with a case‑insensitive regex, builds an equivalent XLOOKUP formula using INDEX for lookup and return arrays, substitutes the old formula, and saves the updated file.
class Program
{
    static void Main()
    {
        // Load the existing workbook (lifecycle rule)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Iterate through all cells that contain formulas
            foreach (Cell cell in cells)
            {
                if (cell.IsFormula && cell.Formula.IndexOf("HLOOKUP", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // Capture the parts of the HLOOKUP formula:
                    // HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])
                    string oldFormula = cell.Formula;

                    Match match = Regex.Match(
                        oldFormula,
                        @"HLOOKUP\(\s*(?<lookup>[^,]+)\s*,\s*(?<table>[^,]+)\s*,\s*(?<row>\d+)\s*(?:,\s*(?<range>[^)]+))?\s*\)",
                        RegexOptions.IgnoreCase);

                    if (match.Success)
                    {
                        string lookupValue = match.Groups["lookup"].Value.Trim();
                        string tableArray = match.Groups["table"].Value.Trim();
                        string rowIndex = match.Groups["row"].Value.Trim();

                        // Build the equivalent XLOOKUP formula:
                        // XLOOKUP(lookup_value,
                        //         INDEX(table_array,1,0),          // first row as lookup array
                        //         INDEX(table_array,row_index,0), // desired row as return array
                        //         "#N/A",                         // not found result
                        //         0)                              // exact match
                        string xlookupFormula = $"XLOOKUP({lookupValue},INDEX({tableArray},1,0),INDEX({tableArray},{rowIndex},0),\"#N/A\",0)";

                        // Replace the HLOOKUP formula with the new XLOOKUP formula
                        cell.Formula = xlookupFormula;
                    }
                }
            }
        }

        // Save the modified workbook (lifecycle rule)
        workbook.Save("output.xlsx");
    }
}
