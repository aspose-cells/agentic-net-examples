// Title: Bulk replace exact‑match VLOOKUP formulas with XLOOKUP in Excel files using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that scans every worksheet, finds VLOOKUP formulas where the fourth argument is FALSE or 0, and rewrites them as XLOOKUP expressions. | Create a helper method that parses a VLOOKUP formula string, verifies it uses exact match, and returns the equivalent XLOOKUP formula built with INDEX for the lookup and return arrays. | Show how to save the updated workbook after converting all matching formulas, preserving original formatting and calculation settings.
// Common Searches: Aspose.Cells C# replace VLOOKUP FALSE with XLOOKUP in all sheets | how to convert exact match VLOOKUP to XLOOKUP programmatically .NET | bulk update Excel formulas from VLOOKUP to XLOOKUP using Aspose.Cells | detect and change VLOOKUP formulas in a workbook with C# Aspose.Cells
// Tags: VLOOKUP to XLOOKUP conversion Aspose.Cells | exact-match VLOOKUP detection C# | bulk formula update Excel .NET | XLOOKUP formula generation Aspose.Cells | Excel workbook cell iteration Aspose.Cells

using System;
using Aspose.Cells;

// The example loads an Excel workbook, iterates through each worksheet and cell, identifies VLOOKUP formulas that use FALSE or 0 for exact matching, converts them to equivalent XLOOKUP formulas using INDEX for lookup and return arrays, and saves the modified workbook.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all cells that contain formulas
            foreach (Cell cell in sheet.Cells)
            {
                if (cell.IsFormula)
                {
                    string formula = cell.Formula;

                    // Identify VLOOKUP formulas that use exact match (FALSE or 0 as the fourth argument)
                    if (IsExactMatchVlookup(formula))
                    {
                        // Convert the VLOOKUP formula to an equivalent XLOOKUP formula
                        string newFormula = ConvertVlookupToXlookup(formula);
                        cell.Formula = newFormula;
                    }
                }
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }

    // Checks whether a formula is a VLOOKUP with an exact‑match flag
    static bool IsExactMatchVlookup(string formula)
    {
        if (string.IsNullOrEmpty(formula) || !formula.StartsWith("="))
            return false;

        int vlookupPos = formula.IndexOf("VLOOKUP", StringComparison.OrdinalIgnoreCase);
        if (vlookupPos < 0)
            return false;

        int openParen = formula.IndexOf('(', vlookupPos);
        int closeParen = formula.LastIndexOf(')');
        if (openParen < 0 || closeParen < 0 || closeParen <= openParen)
            return false;

        // Extract the argument list inside VLOOKUP(...)
        string args = formula.Substring(openParen + 1, closeParen - openParen - 1);

        // Simple split on commas (works for most cases without nested commas)
        string[] parts = args.Split(',');
        if (parts.Length < 4)
            return false;

        string fourthArg = parts[3].Trim().TrimEnd(')');
        return string.Equals(fourthArg, "FALSE", StringComparison.OrdinalIgnoreCase) ||
               fourthArg == "0";
    }

    // Converts a VLOOKUP(lookup_value, table_array, col_index_num, FALSE) to XLOOKUP
    static string ConvertVlookupToXlookup(string formula)
    {
        int vlookupPos = formula.IndexOf("VLOOKUP", StringComparison.OrdinalIgnoreCase);
        int openParen = formula.IndexOf('(', vlookupPos);
        int closeParen = formula.LastIndexOf(')');

        // Extract arguments
        string args = formula.Substring(openParen + 1, closeParen - openParen - 1);
        string[] parts = args.Split(',');

        string lookupValue = parts[0].Trim();
        string tableArray = parts[1].Trim();
        string colIndexStr = parts[2].Trim();

        // Build XLOOKUP:
        // XLOOKUP(lookup_value, INDEX(table_array,0,1), INDEX(table_array,0,col_index_num))
        string xlookup = $"=XLOOKUP({lookupValue},INDEX({tableArray},0,1),INDEX({tableArray},0,{colIndexStr}))";

        return xlookup;
    }
}
